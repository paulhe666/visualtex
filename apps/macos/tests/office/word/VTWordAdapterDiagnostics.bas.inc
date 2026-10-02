' Test-only declarations and procedures. See scripts/word_vba_source.mjs.
Private VT_NATIVE_PARITY_DIAGNOSTIC As Boolean
Private VT_WORD_TYPOGRAPHY_REGRESSION_RUNNING As Boolean
Private VT_WORD_COPY_CONVERSION_REGRESSION_DOCUMENT As Document

Public Sub VisualTeX_RunEditOpenAndTabsRegression()
    Dim shape As InlineShape
    Dim numberRange As Range
    Dim characterRange As Range
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim alignment As Long
    Dim report As String
    Dim sample As String
    Dim encoded As String
    Dim started As Single
    Dim size As Variant
    Dim count As Long
    Dim tabCount As Long
    Dim tabPositions As String
    Dim tabStop As TabStop
    Dim failure As Long
    Dim description As String
    On Error GoTo Failed
    report = "revision=" & VT_WORD_SOURCE_REVISION & vbLf
    If Not VTProtocolSelfTest() Then Err.Raise 5, , "Protocol self-test failed."
    For Each size In Array(16384, 131072)
        sample = String$(CLng(size), "x") & ChrW(20013) & ChrW(960) & ChrW(-10179) & ChrW(-8704)
        started = Timer
        encoded = VTBase64UrlEncodeUtf8(sample)
        report = report & "encode-" & CStr(size) & "-ms=" & _
            CStr(CLng(1000# * VTTimerElapsedSeconds(started, Timer))) & vbLf
        started = Timer
        If VTBase64UrlDecodeUtf8(encoded) <> sample Then Err.Raise 5, , "Large Unicode payload did not round-trip."
        report = report & "decode-" & CStr(size) & "-ms=" & _
            CStr(CLng(1000# * VTTimerElapsedSeconds(started, Timer))) & vbLf
    Next size
    VTBeginWordInternalMutation
    For Each shape In ActiveDocument.InlineShapes
        If VTTryWordImageTypographyReference(shape, formulaId, displayMode, numbered) Then
            If displayMode = "block" And numbered Then
                Set numberRange = VTStaticImageEquationNumberRange(shape.Range, formulaId)
                If numberRange Is Nothing Then Err.Raise 5, , "Numbered image lost its visible number."
                alignment = shape.Range.ParagraphFormat.BaseLineAlignment
                tabCount = shape.Range.ParagraphFormat.TabStops.Count
                tabPositions = ""
                For Each tabStop In shape.Range.ParagraphFormat.TabStops
                    tabPositions = tabPositions & CStr(tabStop.Position) & ":" & CStr(tabStop.Alignment) & ";"
                Next tabStop
                VTApplyNumberedImageFormulaVerticalAlignment shape, numberRange, _
                    VTVisibleEquationNumberFontSize(ActiveDocument), "Tab regression"
                If shape.Range.ParagraphFormat.BaseLineAlignment <> alignment Then _
                    Err.Raise 5, , "The user's Chinese typography alignment changed."
                If shape.Range.ParagraphFormat.TabStops.Count <> tabCount Then _
                    Err.Raise 5, , "Horizontal tab stops changed."
                For Each tabStop In shape.Range.ParagraphFormat.TabStops
                    If InStr(1, tabPositions, CStr(tabStop.Position) & ":" & CStr(tabStop.Alignment) & ";", vbBinaryCompare) = 0 Then _
                        Err.Raise 5, , "Horizontal tab positions changed."
                Next tabStop
                For Each characterRange In shape.Range.Paragraphs(1).Range.Characters
                    If characterRange.Text = vbTab Or characterRange.Text = vbCr Then
                        If characterRange.Font.Position <> 0 Or _
                           characterRange.Font.Size <> numberRange.Font.Size Or _
                           characterRange.Font.NameAscii <> numberRange.Font.NameAscii Then _
                            Err.Raise 5, , "A tab or paragraph mark still has inherited lowering."
                    End If
                Next characterRange
                count = count + 1
            End If
        End If
    Next shape
    report = report & "numbered-images=" & CStr(count) & vbLf & "status=PASS" & vbLf
Finished:
    On Error Resume Next
    VTEndWordInternalMutation
    VTWriteTextAtomic VTApplicationSupportRoot() & "/Tests/edit-open-tabs-regression.txt", _
        report & "error=" & CStr(failure) & ":" & description
    Exit Sub
Failed:
    failure = Err.Number
    description = Err.Description
    Resume Finished
End Sub

Public Function VTProbeInlineShapeRangeFontSizeBehavior() As Boolean
    Dim probeDocument As Document
    Dim probeRange As Range
    Dim probeShape As InlineShape
    Dim firstAppliedSize As Double
    Dim secondAppliedSize As Double

    On Error GoTo ProbeFinished
    If Not VTPathFileExists(VTPlaceholderImagePath()) Then Exit Function
    Set probeDocument = Documents.Add(Visible:=False)
    Set probeRange = probeDocument.Content
    probeRange.Collapse wdCollapseStart
    Set probeShape = probeDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Range:=probeRange)
    probeShape.Range.Font.Size = 10.5
    firstAppliedSize = probeShape.Range.Font.Size
    probeShape.Range.Font.Size = 18#
    secondAppliedSize = probeShape.Range.Font.Size
    VTProbeInlineShapeRangeFontSizeBehavior = _
        Abs(firstAppliedSize - 10.5) <= 0.05 And _
        Abs(secondAppliedSize - 18#) <= 0.05

ProbeFinished:
    On Error Resume Next
    If Not probeDocument Is Nothing Then
        probeDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
End Function

Public Sub VisualTeX_ProbeImageFormulaFontSize()
    If VTProbeInlineShapeRangeFontSizeBehavior() Then
        VTShowInformation _
            "This Mac Word build persists InlineShape.Range.Font.Size. " & _
            "The native Home font-size box can drive VisualTeX image formulas."
    Else
        VTShowInformation _
            "This Mac Word build does not reliably persist image Range.Font.Size. " & _
            "Use the VisualTeX 图片公式字号 drop-down as the supported fallback."
    End If
End Sub

Public Sub VisualTeX_AssertWordHostSelfTest()
    If Not VTProtocolSelfTest() Then
        Err.Raise vbObjectError + 7480, "VisualTeX", _
            "The VisualTeX Word protocol self-test failed."
    End If
    If Not VTWordSourceSelfTest() Then
        Err.Raise vbObjectError + 7481, "VisualTeX", _
            "The VisualTeX Word source self-test failed."
    End If
    VTInitializeWordEvents
End Sub

Public Sub VisualTeX_DumpActiveDocumentOpenXmlForRegression()
    Dim outputPath As String

    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7996, "VisualTeX regression", _
            "No active Word document is available for OpenXML inspection."
    End If
    outputPath = VTApplicationSupportRoot() & _
        "/Tests/document-import-openxml.xml"
    VTWriteTextAtomic outputPath, ActiveDocument.Content.WordOpenXML
End Sub

Public Sub VisualTeX_DumpWordImageTypographyForRegression()
    Dim shape As InlineShape
    Dim result As String
    Dim textSize As Double
    Dim textFont As String
    Dim textSample As String
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim fontSize As Double
    Dim referenceWidth As Double
    Dim referenceHeight As Double
    Dim referenceBaseline As Double
    Dim observedSize As Double
    Dim metrics As String
    Dim request As String
    Dim identified As Boolean
    Dim expectedPosition As Long
    Dim calculationError As String
    Dim fingerprint As String
    Dim inkTop As Double
    Dim inkBottom As Double
    Dim anchorRatio As Double
    Dim ascent As Double
    Dim descent As Double
    Dim asianText As Boolean
    Dim primaryLetter As Boolean
    If Documents.Count = 0 Then Exit Sub
    On Error Resume Next
    fingerprint = VTImageTypographyFingerprint(ActiveDocument.Content)
    result = "HOST" & vbTab & VT_WORD_SOURCE_REVISION & vbTab & _
        CStr(VTWordInternalMutationActive()) & vbTab & CStr(ActiveDocument.ReadOnly) & vbTab & _
        CStr(ActiveDocument.ProtectionType) & vbTab & CStr(Err.Number) & ":" & Err.Description & vbLf
    Err.Clear
    On Error GoTo 0
    For Each shape In ActiveDocument.InlineShapes
        VTWordInlineTextContext shape, textSize, textFont, textSample
        identified = VTTryParseFormulaReference(shape.Title, formulaId, displayMode, numbered)
        metrics = ""
        If identified And VTTryReadWordImageScaleState(ActiveDocument, formulaId, _
           fontSize, referenceWidth, referenceHeight, referenceBaseline, observedSize) Then
            request = "{""bodyFont"":" & VTJsonString(textFont) & _
                ",""bodySample"":" & VTJsonString(textSample) & _
                ",""metadata"":" & VTJsonString(shape.AlternativeText) & "}"
            metrics = VTFileBridgeCall("ReadVisualTeXWordAlignmentMetrics", request)
        End If
        On Error Resume Next
        Err.Clear
        expectedPosition = VTExpectedWordInlineImageBaseline(shape, referenceHeight, referenceBaseline)
        calculationError = CStr(Err.Number) & ":" & Err.Description
        Err.Clear
        VTWordInlineAlignmentMetrics shape, referenceHeight, textSize, inkTop, inkBottom, _
            anchorRatio, ascent, descent, asianText, primaryLetter
        On Error GoTo 0
        result = result & shape.Title & vbTab & CStr(identified) & vbTab & _
            CStr(shape.Range.ParagraphFormat.BaseLineAlignment) & vbTab & _
            CStr(textSize) & vbTab & textFont & vbTab & textSample & vbTab & _
            CStr(shape.Range.Font.Position) & vbTab & CStr(shape.Width) & vbTab & _
            CStr(shape.Height) & vbTab & metrics & vbTab & CStr(expectedPosition) & vbTab & _
            calculationError & vbTab & CStr(inkTop) & "|" & CStr(inkBottom) & "|" & _
            CStr(anchorRatio) & vbLf
    Next shape
    VTWriteTextAtomic VTApplicationSupportRoot() & _
        "/Tests/word-image-typography-probe.tsv", result
End Sub

Public Sub VisualTeX_ExportActiveDocumentPdfForRegression()
    Dim requestPath As String
    Dim statusPath As String
    Dim outputPath As String
    Dim runtimeRoot As String
    Dim scratchRoot As String
    Dim homeBoundary As Long
    Dim exportStage As String
    Dim exportErrorNumber As Long
    Dim exportErrorDescription As String

    runtimeRoot = VTApplicationSupportRoot()
    statusPath = _
        runtimeRoot & _
        "/document-import-regression-pdf-status.txt"
    On Error GoTo Failed
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7590, "VisualTeX", _
            "The PDF regression export requires an active Word document."
    End If
    requestPath = _
        runtimeRoot & _
        "/document-import-regression-pdf-path.txt"
    outputPath = Trim$(VTReadText(requestPath, 4096))
    homeBoundary = InStr( _
        1, runtimeRoot, "/Library/Application Scripts/", vbBinaryCompare)
    If homeBoundary <= 1 Then
        Err.Raise vbObjectError + 7590, "VisualTeX", _
            "The PDF regression export could not resolve the user home path."
    End If
    scratchRoot = _
        Left$(runtimeRoot, homeBoundary - 1) & _
        "/Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch/"
    If Len(outputPath) <= Len(scratchRoot) Or _
       Left$(outputPath, Len(scratchRoot)) <> scratchRoot Or _
       InStr(outputPath, "..") > 0 Or _
       InStr(outputPath, vbCr) > 0 Or _
       InStr(outputPath, vbLf) > 0 Or _
       InStr(outputPath, Chr$(0)) > 0 Then
        Err.Raise vbObjectError + 7590, "VisualTeX", _
            "The PDF regression export rejected an unsafe output path."
    End If

    exportStage = "ExportAsFixedFormat"
    On Error Resume Next
    ActiveDocument.ExportAsFixedFormat _
        OutputFileName:=outputPath, _
        ExportFormat:=wdExportFormatPDF
    exportErrorNumber = Err.Number
    exportErrorDescription = Err.Description
    Err.Clear
    On Error GoTo Failed
    If exportErrorNumber <> 0 Then
        exportStage = "SaveAs2"
        ActiveDocument.SaveAs2 _
            FileName:=outputPath, _
            FileFormat:=wdFormatPDF, _
            AddToRecentFiles:=False
    End If
    VTWriteTextAtomic statusPath, "ok|" & exportStage
    Exit Sub

Failed:
    exportErrorNumber = Err.Number
    exportErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        statusPath, _
        "error|" & CStr(exportErrorNumber) & "|" & _
        exportErrorDescription & "|stage=" & exportStage & _
        "|root=" & runtimeRoot & "|output=" & outputPath
    On Error GoTo 0
End Sub

Public Sub VisualTeX_RunSelectedImageEditRecoveryRegression()
    Dim statusPath As String
    Dim selectedShape As InlineShape
    Dim documentObject As Document
    Dim encodedMetadata As String
    Dim storedMetadata As String
    Dim formulaReference As String
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    statusPath = _
        VTApplicationSupportRoot() & _
        "/document-import-regression-image-edit-status.txt"
    On Error GoTo Failed
    If Documents.Count = 0 Or Selection.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The image edit recovery regression requires one selected inline shape."
    End If
    Set selectedShape = Selection.InlineShapes(1)
    Set documentObject = selectedShape.Range.Document
    encodedMetadata = selectedShape.AlternativeText
    If Not VTIsEncodedMetadata(encodedMetadata) Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The selected regression image is missing VisualTeX metadata."
    End If

    formulaReference = selectedShape.Title
    If Not VTTryParseFormulaReference( _
       formulaReference, formulaId, displayMode, numbered) Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The selected regression image is missing its compact reference."
    End If

    ' Title-only recovery: Variables must restore AlternativeText.
    selectedShape.AlternativeText = ""
    If Not VTIsVisualTeXInlineShape(selectedShape) Or _
       StrComp(selectedShape.AlternativeText, encodedMetadata, _
           vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "Word Variables did not restore the image AlternativeText."
    End If

    ' Word SVG round-trips can swap the two accessibility properties.
    selectedShape.Title = encodedMetadata
    selectedShape.AlternativeText = formulaReference
    If Not VTIsVisualTeXInlineShape(selectedShape) Or _
       StrComp(selectedShape.Title, formulaReference, _
           vbBinaryCompare) <> 0 Or _
       StrComp(selectedShape.AlternativeText, encodedMetadata, _
           vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "Double-click recognition did not repair swapped image metadata."
    End If

    ' Metadata-only recovery: the unique Variables payload must restore Title.
    selectedShape.Title = ""
    If Len(selectedShape.Title) <> 0 Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "Word did not clear the image formula Title for the regression."
    End If
    If Not VTHandleWordBeforeDoubleClick(Selection) Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The real double-click handler did not recognize the image formula."
    End If

    formulaReference = selectedShape.Title
    If Not VTTryParseFormulaReference( _
       formulaReference, formulaId, displayMode, numbered) Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "Double-click editing did not restore the image formula Title."
    End If
    If Not VTTryReadWordMetadataPayload( _
       documentObject, formulaId, storedMetadata) Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The restored image formula UUID has no Word metadata payload."
    End If
    If StrComp(storedMetadata, encodedMetadata, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The restored image formula UUID does not match AlternativeText."
    End If
    VTWriteTextAtomic statusPath, "ok|" & formulaReference
    Exit Sub

Failed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        statusPath, _
        "error|" & CStr(regressionErrorNumber) & "|" & _
        regressionErrorDescription
    On Error GoTo 0
End Sub

Private Sub VTAssertWordImageMacroButtonFixture( _
    ByVal documentObject As Document, _
    ByVal expectedTitle As String, _
    ByVal expectedMetadata As String, _
    ByVal expectedWidth As Single, _
    ByVal expectedHeight As Single, _
    ByVal expectedBaseline As Long, _
    ByVal stageName As String)

    Dim formulaShape As InlineShape
    Dim macroField As Field

    If documentObject Is Nothing Or _
       documentObject.InlineShapes.Count <> 1 Or _
       documentObject.Fields.Count <> 1 Then
        Err.Raise vbObjectError + 7594, "VisualTeX regression", _
            stageName & ": the document does not contain one image field."
    End If
    Set formulaShape = documentObject.InlineShapes(1)
    Set macroField = _
        VTVisualTeXImageMacroButtonFieldForShape(formulaShape)
    If macroField Is Nothing Or macroField.Type <> 51 Or _
       Not VTMacroButtonResultHasOnlyImage(macroField) Then
        Err.Raise vbObjectError + 7594, "VisualTeX regression", _
            stageName & ": the field is not one clean type-51 MacroButton."
    End If
    If StrComp(formulaShape.Title, expectedTitle, vbBinaryCompare) <> 0 Or _
       StrComp(formulaShape.AlternativeText, expectedMetadata, _
           vbBinaryCompare) <> 0 Or _
       Abs(formulaShape.Width - expectedWidth) > 0.1 Or _
       Abs(formulaShape.Height - expectedHeight) > 0.1 Or _
       formulaShape.Range.Font.Position <> expectedBaseline Then
        Err.Raise vbObjectError + 7594, "VisualTeX regression", _
            stageName & ": MacroButton persistence changed image state."
    End If
End Sub

Public Sub VisualTeX_RunWordImageMacroButtonRegression()
    Const formulaId As String = _
        "66666666-6666-4666-8666-666666666666"

    Dim sourceDocument As Document
    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim insertionRange As Range
    Dim encodedMetadata As String
    Dim formulaReference As String
    Dim fixturePath As String
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    If Documents.Count > 0 Then Set sourceDocument = ActiveDocument
    fixturePath = VTApplicationSupportRoot() & _
        "/Tests/word-image-macrobutton-regression.docx"
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-image-macrobutton-regression-result.txt"
    encodedMetadata = VT_METADATA_PREFIX & "e30"
    formulaReference = VTFormulaReference(formulaId, "inline", False)
    On Error Resume Next
    Kill fixturePath
    Err.Clear
    On Error GoTo RegressionFailed

    regressionStage = "create"
    Set testDocument = Documents.Add(Visible:=False)
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    Set formulaShape = testDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    formulaShape.LockAspectRatio = msoFalse
    formulaShape.Width = 73.5!
    formulaShape.Height = 26.25!
    formulaShape.LockAspectRatio = msoTrue
    formulaShape.Range.Font.Position = -2
    formulaShape.Title = formulaReference
    formulaShape.AlternativeText = encodedMetadata
    VTSetWordMetadataPayload testDocument, formulaId, encodedMetadata
    VTSetWordFormulaFormat testDocument, formulaId, "inline", False

    regressionStage = "wrap"
    Set formulaShape = VTEnsureVisualTeXImageMacroButton(formulaShape)
    VTAssertWordImageMacroButtonFixture _
        testDocument, formulaReference, encodedMetadata, _
        73.5!, 26.25!, -2, "initial wrap"

    regressionStage = "idempotent-wrap"
    Set formulaShape = VTEnsureVisualTeXImageMacroButton(formulaShape)
    VTAssertWordImageMacroButtonFixture _
        testDocument, formulaReference, encodedMetadata, _
        73.5!, 26.25!, -2, "idempotent wrap"

    regressionStage = "save"
    testDocument.SaveAs2 _
        FileName:=fixturePath, FileFormat:=wdFormatXMLDocument, _
        AddToRecentFiles:=False
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing

    regressionStage = "reopen"
    Set testDocument = Documents.Open( _
        FileName:=fixturePath, ConfirmConversions:=False, _
        ReadOnly:=False, AddToRecentFiles:=False, Visible:=False)
    VTAssertWordImageMacroButtonFixture _
        testDocument, formulaReference, encodedMetadata, _
        73.5!, 26.25!, -2, "save/reopen"
    VTWriteTextAtomic _
        resultPath, _
        "PASS" & vbLf & _
        "revision=" & VT_WORD_SOURCE_REVISION & vbLf & _
        "fieldType=" & CStr(testDocument.Fields(1).Type) & vbLf & _
        "inlineShapes=" & _
            CStr(testDocument.Fields(1).Result.InlineShapes.Count) & vbLf
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    On Error GoTo 0
    Err.Raise regressionErrorNumber, "VisualTeX MacroButton regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordCopyIdentityRegression()
    Const formulaId As String = _
        "67676767-6767-4767-8767-676767676767"

    Dim sourceDocument As Document
    Dim testDocument As Document
    Dim pastedDocument As Document
    Dim sourceShape As InlineShape
    Dim copiedShape As InlineShape
    Dim crossDocumentShape As InlineShape
    Dim insertionRange As Range
    Dim pasteRange As Range
    Dim ownerShape As InlineShape
    Dim editBookmark As Bookmark
    Dim encodedMetadata As String
    Dim formulaReference As String
    Dim cacheFileName As String
    Dim vectorDocumentPath As String
    Dim fallbackImagePath As String
    Dim sessionId As String
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    If Documents.Count > 0 Then Set sourceDocument = ActiveDocument
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-copy-identity-regression-result.txt"
    encodedMetadata = VT_METADATA_PREFIX & "e30"
    formulaReference = VTFormulaReference(formulaId, "inline", False)

    regressionStage = "create-source"
    cacheFileName = Dir$( _
        VTApplicationSupportRoot() & "/ImageDocuments/*.docx")
    If Len(cacheFileName) = 0 Then
        Err.Raise vbObjectError + 7604, "VisualTeX regression", _
            "No real VisualTeX SVG staging DOCX is available for the copy fixture."
    End If
    vectorDocumentPath = _
        VTApplicationSupportRoot() & "/ImageDocuments/" & cacheFileName
    fallbackImagePath = Left$( _
        vectorDocumentPath, Len(vectorDocumentPath) - Len(".docx")) & ".png"
    If Not VTPathFileExists(fallbackImagePath) Then
        Err.Raise vbObjectError + 7604, "VisualTeX regression", _
            "The VisualTeX SVG staging cache has no PNG compatibility preview."
    End If
    Set testDocument = Documents.Add(Visible:=False)
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    Set sourceShape = VTAddWordFormulaPicture( _
        testDocument, insertionRange, vectorDocumentPath, fallbackImagePath)
    sourceShape.Title = formulaReference
    sourceShape.AlternativeText = encodedMetadata
    VTSetWordMetadataPayload testDocument, formulaId, encodedMetadata
    VTSetWordFormulaFormat testDocument, formulaId, "inline", False
    VTSetWordImageOwnerBookmark testDocument, sourceShape, formulaId

    regressionStage = "copy-paste"
    sourceShape.Range.Copy
    testDocument.Content.InsertAfter vbCr
    Set pasteRange = testDocument.Range( _
        Start:=testDocument.Content.End - 1, _
        End:=testDocument.Content.End - 1)
    pasteRange.Paste
    If testDocument.InlineShapes.Count <> 2 Then
        Err.Raise vbObjectError + 7604, "VisualTeX regression", _
            "Word did not create exactly one copied formula image."
    End If
    Set sourceShape = testDocument.InlineShapes(1)
    Set copiedShape = testDocument.InlineShapes(2)
    If copiedShape.AlternativeText <> encodedMetadata Then
        Err.Raise vbObjectError + 7604, "VisualTeX regression", _
            "Word did not preserve copied formula metadata on the pasted image."
    End If
    Set ownerShape = VTWordImageOwnerShape(testDocument, formulaId)
    If ownerShape Is Nothing Or _
       ownerShape.Range.Start <> sourceShape.Range.Start Then
        Err.Raise vbObjectError + 7604, "VisualTeX regression", _
            "Copy/paste moved the formula owner bookmark away from the source."
    End If
    If VTShouldForkWordImageFormulaCopy( _
       testDocument, sourceShape, formulaId, encodedMetadata) Then
        Err.Raise vbObjectError + 7604, "VisualTeX regression", _
            "The original Word formula was incorrectly classified as a copy."
    End If
    If Not VTShouldForkWordImageFormulaCopy( _
       testDocument, copiedShape, formulaId, encodedMetadata) Then
        Err.Raise vbObjectError + 7604, "VisualTeX regression", _
            "The copied Word formula did not receive copy-isolation semantics."
    End If

    regressionStage = "edit-anchor"
    sessionId = VTNewUuidV4()
    VTAddWordEditBookmark copiedShape, sessionId
    If Not testDocument.Bookmarks.Exists(VTWordEditBookmarkName(sessionId)) Then
        Err.Raise vbObjectError + 7604, "VisualTeX regression", _
            "The copied formula edit bookmark was not created."
    End If
    Set editBookmark = _
        testDocument.Bookmarks(VTWordEditBookmarkName(sessionId))
    If editBookmark.Range.InlineShapes.Count <> 1 Or _
       editBookmark.Range.InlineShapes(1).Range.Start <> copiedShape.Range.Start Then
        Err.Raise vbObjectError + 7604, "VisualTeX regression", _
            "The copied formula edit bookmark resolved the wrong physical object."
    End If
    VTDeleteWordEditBookmark testDocument, sessionId

    regressionStage = "cross-document-copy"
    sourceShape.Range.Copy
    Set pastedDocument = Documents.Add(Visible:=False)
    Set pasteRange = pastedDocument.Range(Start:=0, End:=0)
    pasteRange.Paste
    If pastedDocument.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7604, "VisualTeX regression", _
            "Word did not paste one formula into the second document."
    End If
    Set crossDocumentShape = pastedDocument.InlineShapes(1)
    If Not VTShouldForkWordImageFormulaCopy( _
       pastedDocument, crossDocumentShape, formulaId, encodedMetadata) Then
        Err.Raise vbObjectError + 7604, "VisualTeX regression", _
            "A cross-document Word formula copy retained the source identity."
    End If

    VTWriteTextAtomic _
        resultPath, _
        "PASS" & vbLf & _
        "revision=" & VT_WORD_SOURCE_REVISION & vbLf & _
        "sourceStart=" & CStr(sourceShape.Range.Start) & vbLf & _
        "copyStart=" & CStr(copiedShape.Range.Start) & vbLf & _
        "ownerBookmark=" & VTWordImageOwnerBookmarkName(formulaId) & vbLf
    pastedDocument.Close SaveChanges:=wdDoNotSaveChanges
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    If Not pastedDocument Is Nothing Then
        pastedDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    On Error GoTo 0
    Err.Raise regressionErrorNumber, "VisualTeX Word copy identity regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordNumberedCopyIdentityRegression()
    Const formulaId As String = _
        "61616161-6161-4616-8616-616161616161"

    Dim sourceDocument As Document
    Dim testDocument As Document
    Dim sourceShape As InlineShape
    Dim copiedShape As InlineShape
    Dim canonicalShape As InlineShape
    Dim canonicalRange As Range
    Dim insertionRange As Range
    Dim pasteRange As Range
    Dim encodedMetadata As String
    Dim ownerBookmarkName As String
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    If Documents.Count > 0 Then Set sourceDocument = ActiveDocument
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-numbered-copy-identity-regression-result.txt"
    encodedMetadata = VT_METADATA_PREFIX & "e30"
    ownerBookmarkName = VTWordImageOwnerBookmarkName(formulaId)

    regressionStage = "create-numbered-source"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    Set sourceShape = testDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    sourceShape.LockAspectRatio = msoFalse
    sourceShape.Width = 72!
    sourceShape.Height = 24!
    sourceShape.LockAspectRatio = msoTrue
    sourceShape.Title = VTFormulaReference(formulaId, "block", True)
    sourceShape.AlternativeText = encodedMetadata
    VTSetWordMetadataPayload testDocument, formulaId, encodedMetadata
    VTSetWordFormulaFormat testDocument, formulaId, "block", True
    Set canonicalRange = VTInsertEquationNumber( _
        sourceShape, formulaId, "numbered copy identity regression")
    Set canonicalRange = VTNumberedFormulaRangeForId(testDocument, formulaId)
    If canonicalRange Is Nothing Or canonicalRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "The numbered source formula did not expose one canonical image."
    End If
    Set sourceShape = canonicalRange.InlineShapes(1)
    VTSetWordImageOwnerBookmark testDocument, sourceShape, formulaId

    regressionStage = "copy-numbered-image"
    sourceShape.Range.Copy
    testDocument.Content.InsertAfter vbCr
    Set pasteRange = testDocument.Range( _
        Start:=testDocument.Content.End - 1, _
        End:=testDocument.Content.End - 1)
    pasteRange.Paste
    If testDocument.InlineShapes.Count <> 2 Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "Word did not create exactly one copied numbered formula image."
    End If
    Set sourceShape = testDocument.InlineShapes(1)
    Set copiedShape = testDocument.InlineShapes(2)
    If copiedShape.Title <> sourceShape.Title Or _
       copiedShape.AlternativeText <> encodedMetadata Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "Word did not preserve numbered formula identity metadata on copy."
    End If

    regressionStage = "resolve-numbered-canonical-owner"
    Set canonicalRange = VTNumberedFormulaRangeForId(testDocument, formulaId)
    If canonicalRange Is Nothing Or canonicalRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "The numbered bookmarks no longer resolve one canonical image."
    End If
    Set canonicalShape = canonicalRange.InlineShapes(1)
    If canonicalShape.Range.Start <> sourceShape.Range.Start Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "The numbered bookmarks resolved the copied image instead of the source."
    End If

    regressionStage = "owner-bookmark-missing"
    If testDocument.Bookmarks.Exists(ownerBookmarkName) Then
        testDocument.Bookmarks(ownerBookmarkName).Delete
    End If
    If VTShouldForkWordImageFormulaCopy( _
       testDocument, sourceShape, formulaId, encodedMetadata) Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "A numbered source formula was classified as a copy without VT_I_."
    End If
    If Not VTShouldForkWordImageFormulaCopy( _
       testDocument, copiedShape, formulaId, encodedMetadata) Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "A numbered formula copy retained the source identity without VT_I_."
    End If

    regressionStage = "owner-bookmark-moved-to-copy"
    VTSetWordImageOwnerBookmark testDocument, copiedShape, formulaId
    If VTShouldForkWordImageFormulaCopy( _
       testDocument, sourceShape, formulaId, encodedMetadata) Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "Numbered anchors did not override a moved image-owner bookmark."
    End If
    If Not VTShouldForkWordImageFormulaCopy( _
       testDocument, copiedShape, formulaId, encodedMetadata) Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "The moved image-owner bookmark made the copy canonical."
    End If

    VTWriteTextAtomic _
        resultPath, _
        "PASS" & vbLf & _
        "revision=" & VT_WORD_SOURCE_REVISION & vbLf & _
        "sourceStart=" & CStr(sourceShape.Range.Start) & vbLf & _
        "copyStart=" & CStr(copiedShape.Range.Start) & vbLf & _
        "numberOwner=" & VTEquationNumberBookmarkName(formulaId) & vbLf
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    If regressionErrorNumber = 0 Then regressionErrorNumber = vbObjectError + 7606
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & regressionErrorDescription & vbLf
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word numbered copy identity regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordNativeCopyIdentityProbe()
    Const formulaId As String = _
        "69696969-6969-4969-8969-696969696969"

    Dim sourceDocument As Document
    Dim testDocument As Document
    Dim pastedDocument As Document
    Dim insertionRange As Range
    Dim formulaRange As Range
    Dim sourceMath As OMath
    Dim copiedMath As OMath
    Dim resolvedMath As OMath
    Dim resolvedSourceDocument As Document
    Dim crossCopiedMath As OMath
    Dim crossResolvedMath As OMath
    Dim crossSourceDocument As Document
    Dim candidateBookmark As Bookmark
    Dim bookmarkSummary As String
    Dim encodedMetadata As String
    Dim resolvedMetadata As String
    Dim resolvedFormulaId As String
    Dim resolvedDisplayMode As String
    Dim resolvedNumbered As Boolean
    Dim crossFormulaId As String
    Dim crossMetadata As String
    Dim crossDisplayMode As String
    Dim crossNumbered As Boolean
    Dim editSessionId As String
    Dim editBookmarkName As String
    Dim crossSessionId As String
    Dim crossEditBookmarkName As String
    Dim crossLocalBookmark As Bookmark
    Dim crossLocalBookmarkFound As Boolean
    Dim sourceSignature As String
    Dim copiedSignature As String
    Dim directMetadata As String
    Dim directDisplayMode As String
    Dim directNumbered As Boolean
    Dim directMetadataReadable As Boolean
    Dim directFormatReadable As Boolean
    Dim directSignatureMatch As Boolean
    Dim directSourceMath As OMath
    Dim directLocalBookmark As Bookmark
    Dim directLocalBookmarkFound As Boolean
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    If Documents.Count > 0 Then Set sourceDocument = ActiveDocument
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-native-copy-identity-probe-result.txt"

    regressionStage = "create-native"
    Set testDocument = Documents.Add(Visible:=False)
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    insertionRange.Text = "x+1=2"
    Set formulaRange = testDocument.OMaths.Add(insertionRange)
    If formulaRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "Word did not create exactly one source native equation."
    End If
    Set sourceMath = formulaRange.OMaths(1)
    sourceMath.BuildUp
    sourceMath.Type = wdOMathInline
    VTSetNativeFormulaBookmark testDocument, sourceMath.Range, formulaId
    encodedMetadata = VT_METADATA_PREFIX & "e30"
    VTSetWordMetadataPayload testDocument, formulaId, encodedMetadata
    VTSetWordFormulaFormat testDocument, formulaId, "inline", False
    VTSetWordNativeSignature testDocument, formulaId, sourceMath

    regressionStage = "copy-paste"
    sourceMath.Range.Copy
    testDocument.Content.InsertAfter vbCr
    Set insertionRange = testDocument.Range( _
        Start:=testDocument.Content.End - 1, _
        End:=testDocument.Content.End - 1)
    insertionRange.Paste
    If testDocument.OMaths.Count <> 2 Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "Word did not create exactly one copied native equation."
    End If

    regressionStage = "resolve-copy-source"
    Set copiedMath = testDocument.OMaths(2)
    sourceSignature = VTNativeMathFastSignature( _
        testDocument, formulaId, sourceMath)
    copiedSignature = VTNativeMathFastSignature( _
        testDocument, formulaId, copiedMath)
    Set directSourceMath = VTNativeMathForBookmark( _
        testDocument.Bookmarks(VTNativeFormulaBookmarkName(formulaId)))
    directMetadataReadable = VTTryReadWordMetadataPayload( _
        testDocument, formulaId, directMetadata)
    directFormatReadable = VTTryReadWordFormulaFormat( _
        testDocument, formulaId, directDisplayMode, directNumbered)
    directSignatureMatch = VTWordNativeSignatureMatches( _
        testDocument, formulaId, copiedMath)
    directLocalBookmarkFound = VTTryFindNativeFormulaBookmarkLocally( _
        copiedMath.Range, directLocalBookmark)
    If Not VTTryResolveCopiedNativeFormulaSource( _
       copiedMath.Range, resolvedMath, resolvedSourceDocument, _
       resolvedFormulaId, resolvedMetadata, resolvedDisplayMode, _
       resolvedNumbered) Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "VisualTeX did not recover the copied native equation source. " & _
            "sourceType=" & CStr(sourceMath.Type) & _
            " copiedType=" & CStr(copiedMath.Type) & _
            " sourceSignature=" & sourceSignature & _
            " copiedSignature=" & copiedSignature & _
            " directSource=" & CStr(Not directSourceMath Is Nothing) & _
            " metadataReadable=" & CStr(directMetadataReadable) & _
            " formatReadable=" & CStr(directFormatReadable) & _
            " directDisplayMode=" & directDisplayMode & _
            " signatureMatch=" & CStr(directSignatureMatch) & _
            " localBookmarkFound=" & CStr(directLocalBookmarkFound) & _
            " sourceText=" & Replace$(sourceMath.Range.Text, vbCr, "<CR>") & _
            " copiedText=" & Replace$(copiedMath.Range.Text, vbCr, "<CR>")
    End If
    If resolvedMath Is Nothing Or resolvedSourceDocument Is Nothing Or _
       VTWordDocumentIdentityForDocument(resolvedSourceDocument) <> _
           VTWordDocumentIdentityForDocument(testDocument) Or _
       resolvedFormulaId <> formulaId Or _
       resolvedMetadata <> encodedMetadata Or _
       resolvedDisplayMode <> IIf( _
           copiedMath.Type = wdOMathInline, "inline", "block") Or _
       resolvedNumbered Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "The copied native equation resolved to the wrong source identity. " & _
            "resolvedFormulaId=" & resolvedFormulaId & _
            " resolvedDisplayMode=" & resolvedDisplayMode & _
            " resolvedNumbered=" & CStr(resolvedNumbered) & _
            " metadataMatch=" & CStr(resolvedMetadata = encodedMetadata) & _
            " documentMatch=" & CStr( _
                VTWordDocumentIdentityForDocument(resolvedSourceDocument) = _
                VTWordDocumentIdentityForDocument(testDocument))
    End If

    regressionStage = "anchor-copy-session"
    editSessionId = VTNewUuidV4()
    editBookmarkName = VTWordEditBookmarkName(editSessionId)
    VTAddWordEditRangeBookmark copiedMath.Range, editSessionId
    If Not testDocument.Bookmarks.Exists(editBookmarkName) Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "VisualTeX did not create the copied native Session Bookmark."
    End If
    If testDocument.Bookmarks(editBookmarkName).Range.Start <> _
       copiedMath.Range.Start Or _
       testDocument.Bookmarks(editBookmarkName).Range.End <> _
       copiedMath.Range.End Or _
       testDocument.Bookmarks(editBookmarkName).Range.Start = _
       sourceMath.Range.Start Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "The copied native Session Bookmark does not isolate the pasted equation."
    End If

    regressionStage = "cross-document-copy"
    Set pastedDocument = Documents.Add(Visible:=False)
    sourceMath.Range.Copy
    Set insertionRange = pastedDocument.Range(Start:=0, End:=0)
    insertionRange.Paste
    If pastedDocument.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "Word did not create one cross-document copied native equation."
    End If
    Set crossCopiedMath = pastedDocument.OMaths(1)
    crossLocalBookmarkFound = VTTryFindNativeFormulaBookmarkLocally( _
        crossCopiedMath.Range, crossLocalBookmark)
    If Not VTTryResolveCopiedNativeFormulaSource( _
       crossCopiedMath.Range, crossResolvedMath, crossSourceDocument, _
       crossFormulaId, crossMetadata, crossDisplayMode, crossNumbered) Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "VisualTeX did not recover the cross-document copied native source. " & _
            "localBookmark=" & CStr(crossLocalBookmarkFound) & _
            " bookmarkName=" & IIf( _
                crossLocalBookmarkFound, crossLocalBookmark.Name, "") & _
            " sourceCopySignature=" & VTNativeMathCopySignature(sourceMath) & _
            " copiedCopySignature=" & VTNativeMathCopySignature(crossCopiedMath)
    End If
    If crossResolvedMath Is Nothing Or crossSourceDocument Is Nothing Or _
       VTWordDocumentIdentityForDocument(crossSourceDocument) <> _
           VTWordDocumentIdentityForDocument(testDocument) Or _
       crossFormulaId <> formulaId Or crossMetadata <> encodedMetadata Or _
       crossNumbered Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "The cross-document native copy resolved to the wrong source identity."
    End If
    crossSessionId = VTNewUuidV4()
    crossEditBookmarkName = VTWordEditBookmarkName(crossSessionId)
    VTAddWordEditRangeBookmark crossCopiedMath.Range, crossSessionId
    If Not pastedDocument.Bookmarks.Exists(crossEditBookmarkName) Or _
       pastedDocument.Bookmarks(crossEditBookmarkName).Range.Start <> _
           crossCopiedMath.Range.Start Or _
       pastedDocument.Bookmarks(crossEditBookmarkName).Range.End <> _
           crossCopiedMath.Range.End Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "The cross-document copied native Session Bookmark is not isolated."
    End If

    regressionStage = "inspect-bookmarks"
    For Each candidateBookmark In testDocument.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_NATIVE_BOOKMARK_PREFIX)) = _
           VT_WORD_NATIVE_BOOKMARK_PREFIX Then
            If Len(bookmarkSummary) > 0 Then bookmarkSummary = bookmarkSummary & ";"
            bookmarkSummary = bookmarkSummary & _
                candidateBookmark.Name & "@" & _
                CStr(candidateBookmark.Range.Start) & ":" & _
                CStr(candidateBookmark.Range.End)
        End If
    Next candidateBookmark
    VTWriteTextAtomic _
        resultPath, _
        "PASS" & vbLf & _
        "revision=" & VT_WORD_SOURCE_REVISION & vbLf & _
        "omaths=" & CStr(testDocument.OMaths.Count) & vbLf & _
        "nativeBookmarks=" & bookmarkSummary & vbLf & _
        "resolvedSourceFormulaId=" & resolvedFormulaId & vbLf & _
        "editBookmark=" & editBookmarkName & "@" & _
            CStr(testDocument.Bookmarks(editBookmarkName).Range.Start) & ":" & _
            CStr(testDocument.Bookmarks(editBookmarkName).Range.End) & vbLf & _
        "crossDocumentSourceFormulaId=" & crossFormulaId & vbLf & _
        "crossEditBookmark=" & crossEditBookmarkName & "@" & _
            CStr(pastedDocument.Bookmarks(crossEditBookmarkName).Range.Start) & ":" & _
            CStr(pastedDocument.Bookmarks(crossEditBookmarkName).Range.End) & vbLf & _
        "firstRange=" & CStr(testDocument.OMaths(1).Range.Start) & ":" & _
            CStr(testDocument.OMaths(1).Range.End) & vbLf & _
        "secondRange=" & CStr(testDocument.OMaths(2).Range.Start) & ":" & _
            CStr(testDocument.OMaths(2).Range.End) & vbLf
    pastedDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set pastedDocument = Nothing
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    If Not pastedDocument Is Nothing Then
        pastedDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    On Error GoTo 0
    Err.Raise regressionErrorNumber, "VisualTeX Word native copy probe", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Private Function VTWordCopyConversionRegressionDocument( _
    ByVal expectedIdentity As String) As Document

    Dim candidate As Document

    On Error Resume Next
    If Not VT_WORD_COPY_CONVERSION_REGRESSION_DOCUMENT Is Nothing Then
        If VTWordDocumentIdentityForDocument( _
           VT_WORD_COPY_CONVERSION_REGRESSION_DOCUMENT) = expectedIdentity Then
            Set VTWordCopyConversionRegressionDocument = _
                VT_WORD_COPY_CONVERSION_REGRESSION_DOCUMENT
            Exit Function
        End If
    End If
    Err.Clear
    On Error GoTo 0

    Set candidate = VTFindOpenWordDocumentByIdentity(expectedIdentity)
    If Not candidate Is Nothing Then
        Set VT_WORD_COPY_CONVERSION_REGRESSION_DOCUMENT = candidate
        Set VTWordCopyConversionRegressionDocument = candidate
    End If
End Function

Private Sub VTCloseWordCopyConversionRegressionDocument()
    On Error Resume Next
    If Not VT_WORD_COPY_CONVERSION_REGRESSION_DOCUMENT Is Nothing Then
        VT_WORD_COPY_CONVERSION_REGRESSION_DOCUMENT.Close _
            SaveChanges:=wdDoNotSaveChanges
    End If
    Set VT_WORD_COPY_CONVERSION_REGRESSION_DOCUMENT = Nothing
    On Error GoTo 0
End Sub

Public Sub VisualTeX_RunWordCopiedImageToNativeConversionRegression()
    Const sourceFormulaId As String = _
        "68686868-6868-4868-8868-686868686868"

    Dim sourceDocument As Document
    Dim testDocument As Document
    Dim sourceShape As InlineShape
    Dim copiedShape As InlineShape
    Dim insertionRange As Range
    Dim pasteRange As Range
    Dim encodedMetadata As String
    Dim formulaReference As String
    Dim cacheFileName As String
    Dim vectorDocumentPath As String
    Dim fallbackImagePath As String
    Dim sessionId As String
    Dim inputPath As String
    Dim documentPath As String
    Dim fixturePath As String
    Dim runtimeRoot As String
    Dim homeBoundary As Long
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim launchErrorNumber As Long
    Dim launchErrorDescription As String
    Dim sourceStart As Long
    Dim copyStart As Long

    On Error GoTo RegressionFailed
    If Documents.Count > 0 Then Set sourceDocument = ActiveDocument
    inputPath = VTApplicationSupportRoot() & _
        "/Tests/word-copy-conversion-metadata.txt"
    documentPath = VTApplicationSupportRoot() & _
        "/Tests/word-copy-conversion-document.txt"
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-copy-image-to-native-start.txt"
    runtimeRoot = VTApplicationSupportRoot()
    homeBoundary = InStr( _
        1, runtimeRoot, "/Library/Application Scripts/", vbBinaryCompare)
    If homeBoundary <= 1 Then
        Err.Raise vbObjectError + 7607, "VisualTeX regression", _
            "The copied-conversion fixture could not resolve the user home path."
    End If
    fixturePath = _
        Left$(runtimeRoot, homeBoundary - 1) & _
        "/Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch/" & _
        "word-copy-conversion-fixture.docx"
    encodedMetadata = Trim$(VTReadText(inputPath, 262144))
    If Not VTIsEncodedMetadata(encodedMetadata) Then
        Err.Raise vbObjectError + 7607, "VisualTeX regression", _
            "The copied-conversion metadata fixture is invalid."
    End If

    regressionStage = "create-source-image"
    cacheFileName = Dir$( _
        VTApplicationSupportRoot() & "/ImageDocuments/*.docx")
    If Len(cacheFileName) = 0 Then
        Err.Raise vbObjectError + 7607, "VisualTeX regression", _
            "No real VisualTeX SVG staging DOCX is available for conversion."
    End If
    vectorDocumentPath = _
        VTApplicationSupportRoot() & "/ImageDocuments/" & cacheFileName
    fallbackImagePath = Left$( _
        vectorDocumentPath, Len(vectorDocumentPath) - Len(".docx")) & ".png"
    If Not VTPathFileExists(fallbackImagePath) Then
        Err.Raise vbObjectError + 7607, "VisualTeX regression", _
            "The conversion staging cache has no PNG compatibility preview."
    End If

    On Error Resume Next
    Kill fixturePath
    Err.Clear
    On Error GoTo RegressionFailed
    regressionStage = "create-source-document"
    Set testDocument = Documents.Add
    testDocument.SaveAs2 _
        FileName:=fixturePath, _
        FileFormat:=wdFormatXMLDocument
    Set VT_WORD_COPY_CONVERSION_REGRESSION_DOCUMENT = testDocument
    regressionStage = "insert-source-image"
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    Set sourceShape = VTAddWordFormulaPicture( _
        testDocument, insertionRange, vectorDocumentPath, fallbackImagePath)
    regressionStage = "write-source-image-metadata"
    formulaReference = VTFormulaReference(sourceFormulaId, "inline", False)
    sourceShape.Title = formulaReference
    sourceShape.AlternativeText = encodedMetadata
    sourceShape.Range.Font.Position = 0
    On Error Resume Next
    sourceShape.Range.Font.Size = VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
    Err.Clear
    On Error GoTo RegressionFailed
    VTSetWordMetadataPayload testDocument, sourceFormulaId, encodedMetadata
    VTSetWordFormulaFormat testDocument, sourceFormulaId, "inline", False
    VTSetWordImageScaleState _
        testDocument, sourceFormulaId, VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT, _
        sourceShape.Width, sourceShape.Height, 0#, _
        VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
    regressionStage = "write-source-image-owner"
    VTSetWordImageOwnerBookmark _
        testDocument, sourceShape, sourceFormulaId

    regressionStage = "copy-source-image"
    sourceShape.Range.Copy
    testDocument.Content.InsertAfter vbCr
    Set pasteRange = testDocument.Range( _
        Start:=testDocument.Content.End - 1, _
        End:=testDocument.Content.End - 1)
    pasteRange.Paste
    If testDocument.InlineShapes.Count <> 2 Then
        Err.Raise vbObjectError + 7607, "VisualTeX regression", _
            "Word did not create one copied image for conversion."
    End If
    Set sourceShape = testDocument.InlineShapes(1)
    Set copiedShape = testDocument.InlineShapes(2)
    If Not VTShouldForkWordImageFormulaCopy( _
       testDocument, copiedShape, sourceFormulaId, encodedMetadata) Then
        Err.Raise vbObjectError + 7607, "VisualTeX regression", _
            "The copied image was not classified for conversion isolation."
    End If
    sourceStart = sourceShape.Range.Start
    copyStart = copiedShape.Range.Start

    regressionStage = "launch-image-to-native"
    testDocument.Activate
    copiedShape.Select
    VTWriteTextAtomic _
        documentPath, _
        VTWordDocumentIdentityForDocument(testDocument) & vbLf
    On Error Resume Next
    VTWordEditInlineShape copiedShape, True, "imageToNative", sessionId
    launchErrorNumber = Err.Number
    launchErrorDescription = Err.Description
    Err.Clear
    On Error GoTo RegressionFailed
    If launchErrorNumber <> 0 Then
        Err.Raise launchErrorNumber, _
            "VisualTeX copied image conversion launch", _
            launchErrorDescription
    End If
    If Not VTIsCanonicalUuid(sessionId) Then
        Err.Raise vbObjectError + 7607, "VisualTeX regression", _
            "The copied image conversion did not return its launched Session id."
    End If
    VTWriteTextAtomic _
        resultPath, _
        "START" & vbLf & _
        "revision=" & VT_WORD_SOURCE_REVISION & vbLf & _
        "documentName=" & testDocument.Name & vbLf & _
        "sourceFormulaId=" & sourceFormulaId & vbLf & _
        "sessionId=" & sessionId & vbLf & _
        "sourceStart=" & CStr(sourceStart) & vbLf & _
        "copyStart=" & CStr(copyStart) & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    Set VT_WORD_COPY_CONVERSION_REGRESSION_DOCUMENT = Nothing
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX copied image conversion regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_InspectWordCopiedImageToNativeConversionRegression()
    Const sourceFormulaId As String = _
        "68686868-6868-4868-8868-686868686868"

    Dim testDocument As Document
    Dim ownerShape As InlineShape
    Dim candidateBookmark As Bookmark
    Dim nativeMath As OMath
    Dim newFormulaId As String
    Dim candidateFormulaId As String
    Dim encodedMetadata As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim matchCount As Long
    Dim documentName As String
    Dim documentPath As String
    Dim resultPath As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    documentPath = VTApplicationSupportRoot() & _
        "/Tests/word-copy-conversion-document.txt"
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-copy-image-to-native-result.txt"
    documentName = Replace$(Replace$( _
        Trim$(VTReadText(documentPath, 512)), vbCr, ""), vbLf, "")
    Set testDocument = VTWordCopyConversionRegressionDocument(documentName)
    If testDocument Is Nothing Then
        Err.Raise vbObjectError + 7608, "VisualTeX regression", _
            "The copied image conversion test document is no longer open."
    End If
    If testDocument.InlineShapes.Count <> 1 Or _
       testDocument.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7608, "VisualTeX regression", _
            "Copied image conversion did not leave one source image and one OMath."
    End If
    Set ownerShape = VTWordImageOwnerShape(testDocument, sourceFormulaId)
    If ownerShape Is Nothing Or _
       ownerShape.Range.Start <> testDocument.InlineShapes(1).Range.Start Then
        Err.Raise vbObjectError + 7608, "VisualTeX regression", _
            "Copied image conversion changed the source image owner."
    End If

    For Each candidateBookmark In testDocument.Bookmarks
        candidateFormulaId = ""
        If VTTryFormulaIdFromNativeBookmark( _
           candidateBookmark.Name, candidateFormulaId) Then
            matchCount = matchCount + 1
            newFormulaId = candidateFormulaId
            Set nativeMath = VTNativeMathForBookmark(candidateBookmark)
        End If
    Next candidateBookmark
    If matchCount <> 1 Or Not VTIsCanonicalUuid(newFormulaId) Or _
       newFormulaId = sourceFormulaId Or nativeMath Is Nothing Then
        Err.Raise vbObjectError + 7608, "VisualTeX regression", _
            "Copied image conversion did not create one independent native identity."
    End If
    If Not VTTryReadWordMetadataPayload( _
       testDocument, newFormulaId, encodedMetadata) Or _
       Not VTTryReadWordFormulaFormat( _
           testDocument, newFormulaId, displayMode, numbered) Or _
       Not VTWordNativeSignatureMatches( _
           testDocument, newFormulaId, nativeMath) Then
        Err.Raise vbObjectError + 7608, "VisualTeX regression", _
            "Copied image conversion did not persist independent native state."
    End If
    VTWriteTextAtomic _
        resultPath, _
        "PASS" & vbLf & _
        "sourceFormulaId=" & sourceFormulaId & vbLf & _
        "newFormulaId=" & newFormulaId & vbLf & _
        "displayMode=" & displayMode & vbLf & _
        "numbered=" & CStr(numbered) & vbLf & _
        "metadata=" & encodedMetadata & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX copied image conversion inspection", _
        regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordCopiedNativeToImageConversionRegression()
    Dim testDocument As Document
    Dim sourceBookmark As Bookmark
    Dim candidateBookmark As Bookmark
    Dim sourceMath As OMath
    Dim copiedMath As OMath
    Dim candidateFormulaId As String
    Dim sourceFormulaId As String
    Dim sessionId As String
    Dim insertionRange As Range
    Dim documentName As String
    Dim documentPath As String
    Dim resultPath As String
    Dim copyStart As Long
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    documentPath = VTApplicationSupportRoot() & _
        "/Tests/word-copy-conversion-document.txt"
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-copy-native-to-image-start.txt"
    documentName = Replace$(Replace$( _
        Trim$(VTReadText(documentPath, 512)), vbCr, ""), vbLf, "")
    Set testDocument = VTWordCopyConversionRegressionDocument(documentName)
    If testDocument Is Nothing Then
        Err.Raise vbObjectError + 7609, "VisualTeX regression", _
            "The native-to-image copy test document is no longer open."
    End If
    For Each candidateBookmark In testDocument.Bookmarks
        candidateFormulaId = ""
        If VTTryFormulaIdFromNativeBookmark( _
           candidateBookmark.Name, candidateFormulaId) Then
            If sourceBookmark Is Nothing Then
                Set sourceBookmark = candidateBookmark
                sourceFormulaId = candidateFormulaId
            Else
                Err.Raise vbObjectError + 7609, "VisualTeX regression", _
                    "The native-to-image copy fixture contains multiple native owners."
            End If
        End If
    Next candidateBookmark
    If sourceBookmark Is Nothing Or Not VTIsCanonicalUuid(sourceFormulaId) Then
        Err.Raise vbObjectError + 7609, "VisualTeX regression", _
            "The native-to-image copy fixture has no managed native source."
    End If
    Set sourceMath = VTNativeMathForBookmark(sourceBookmark)
    If sourceMath Is Nothing Then
        Err.Raise vbObjectError + 7609, "VisualTeX regression", _
            "The native-to-image copy source has no OMath."
    End If

    sourceMath.Range.Copy
    testDocument.Content.InsertAfter vbCr
    Set insertionRange = testDocument.Range( _
        Start:=testDocument.Content.End - 1, _
        End:=testDocument.Content.End - 1)
    insertionRange.Paste
    If testDocument.OMaths.Count <> 2 Then
        Err.Raise vbObjectError + 7609, "VisualTeX regression", _
            "Word did not create one copied native equation for conversion."
    End If
    Set copiedMath = testDocument.OMaths(2)
    copyStart = copiedMath.Range.Start
    testDocument.Activate
    copiedMath.Range.Select
    If Not VTWordOpenCopiedNativeSession( _
       copiedMath.Range, False, "nativeToImage", sessionId) Then
        Err.Raise vbObjectError + 7609, "VisualTeX regression", _
            "The copied native conversion did not resolve its source identity."
    End If
    If Not VTIsCanonicalUuid(sessionId) Then
        Err.Raise vbObjectError + 7609, "VisualTeX regression", _
            "The copied native conversion did not return its launched Session id."
    End If
    VTWriteTextAtomic _
        resultPath, _
        "START" & vbLf & _
        "sourceFormulaId=" & sourceFormulaId & vbLf & _
        "sessionId=" & sessionId & vbLf & _
        "copyStart=" & CStr(copyStart) & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX copied native conversion regression", _
        regressionErrorDescription
End Sub

Public Sub VisualTeX_InspectWordCopiedNativeToImageConversionRegression()
    Const originalImageFormulaId As String = _
        "68686868-6868-4868-8868-686868686868"

    Dim testDocument As Document
    Dim candidateBookmark As Bookmark
    Dim candidateShape As InlineShape
    Dim sourceNativeFormulaId As String
    Dim nativeFormulaId As String
    Dim imageFormulaId As String
    Dim candidateFormulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim parsedDisplayMode As String
    Dim parsedNumbered As Boolean
    Dim encodedMetadata As String
    Dim nativeCount As Long
    Dim newImageCount As Long
    Dim ownerShape As InlineShape
    Dim documentName As String
    Dim documentPath As String
    Dim resultPath As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    documentPath = VTApplicationSupportRoot() & _
        "/Tests/word-copy-conversion-document.txt"
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-copy-native-to-image-result.txt"
    documentName = Replace$(Replace$( _
        Trim$(VTReadText(documentPath, 512)), vbCr, ""), vbLf, "")
    Set testDocument = VTWordCopyConversionRegressionDocument(documentName)
    If testDocument Is Nothing Then
        Err.Raise vbObjectError + 7610, "VisualTeX regression", _
            "The copied native conversion test document is no longer open."
    End If
    If testDocument.InlineShapes.Count <> 2 Or _
       testDocument.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7610, "VisualTeX regression", _
            "Copied native conversion did not leave two images and one source OMath."
    End If

    For Each candidateBookmark In testDocument.Bookmarks
        candidateFormulaId = ""
        If VTTryFormulaIdFromNativeBookmark( _
           candidateBookmark.Name, candidateFormulaId) Then
            nativeCount = nativeCount + 1
            sourceNativeFormulaId = candidateFormulaId
        End If
    Next candidateBookmark
    If nativeCount <> 1 Or Not VTIsCanonicalUuid(sourceNativeFormulaId) Then
        Err.Raise vbObjectError + 7610, "VisualTeX regression", _
            "Copied native conversion changed the source native identity."
    End If

    For Each candidateShape In testDocument.InlineShapes
        candidateFormulaId = ""
        parsedDisplayMode = ""
        parsedNumbered = False
        If VTTryParseFormulaReference( _
           candidateShape.Title, candidateFormulaId, _
           parsedDisplayMode, parsedNumbered) Then
            If candidateFormulaId <> originalImageFormulaId Then
                newImageCount = newImageCount + 1
                imageFormulaId = candidateFormulaId
            End If
        End If
    Next candidateShape
    If newImageCount <> 1 Or Not VTIsCanonicalUuid(imageFormulaId) Or _
       imageFormulaId = sourceNativeFormulaId Then
        Err.Raise vbObjectError + 7610, "VisualTeX regression", _
            "Copied native conversion did not create one independent image identity."
    End If
    Set ownerShape = VTWordImageOwnerShape(testDocument, imageFormulaId)
    If ownerShape Is Nothing Then
        Err.Raise vbObjectError + 7610, "VisualTeX regression", _
            "Copied native conversion did not persist the new image owner."
    End If
    If Not VTTryReadWordMetadataPayload( _
       testDocument, imageFormulaId, encodedMetadata) Or _
       Not VTTryReadWordFormulaFormat( _
           testDocument, imageFormulaId, displayMode, numbered) Then
        Err.Raise vbObjectError + 7610, "VisualTeX regression", _
            "Copied native conversion did not persist independent image state."
    End If
    VTWriteTextAtomic _
        resultPath, _
        "PASS" & vbLf & _
        "sourceNativeFormulaId=" & sourceNativeFormulaId & vbLf & _
        "newImageFormulaId=" & imageFormulaId & vbLf & _
        "displayMode=" & displayMode & vbLf & _
        "numbered=" & CStr(numbered) & vbLf & _
        "metadata=" & encodedMetadata & vbLf
    VTCloseWordCopyConversionRegressionDocument
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX copied native conversion inspection", _
        regressionErrorDescription
End Sub

Public Sub VisualTeX_RunDocumentImportFormulaRegression()
    Dim statusPath As String
    Dim documentObject As Document
    Dim candidateBookmark As Bookmark
    Dim candidateMath As OMath
    Dim candidateShape As InlineShape
    Dim formulaRange As Range
    Dim relationRange As Range
    Dim paragraphRange As Range
    Dim formulaText As String
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim nativeFormulaCount As Long
    Dim nativeDisplayCount As Long
    Dim invalidNativeRangeCount As Long
    Dim emptyMathCount As Long
    Dim alignedFormulaCount As Long
    Dim relationCount As Long
    Dim relationIndex As Long
    Dim imageFormulaCount As Long
    Dim imageDisplayCount As Long
    Dim imageMacroButtonCount As Long
    Dim invalidImageMacroButtonCount As Long
    Dim imageFormulaIds As String
    Dim imageMacroField As Field
    Dim relationPosition As Double
    Dim minimumRelationPosition As Double
    Dim maximumRelationPosition As Double
    Dim maximumAlignmentError As Double
    Dim maximumImageSpaceBefore As Double
    Dim maximumImageSpaceAfter As Double
    Dim minimumNativeSpaceBefore As Double
    Dim maximumNativeSpaceBefore As Double
    Dim minimumNativeSpaceAfter As Double
    Dim maximumNativeSpaceAfter As Double
    Dim report As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    statusPath = VTApplicationSupportRoot() & _
        "/document-import-regression-formula-status.txt"
    On Error GoTo Failed
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7592, "VisualTeX", _
            "The document-import formula regression requires an active document."
    End If
    Set documentObject = ActiveDocument
    minimumNativeSpaceBefore = 1000000#
    minimumNativeSpaceAfter = 1000000#

    For Each candidateMath In documentObject.OMaths
        If Not VTOMathHasMeaningfulContent(candidateMath) Then
            emptyMathCount = emptyMathCount + 1
        End If
    Next candidateMath

    For Each candidateBookmark In documentObject.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_NATIVE_BOOKMARK_PREFIX)) = _
           VT_WORD_NATIVE_BOOKMARK_PREFIX Then
            nativeFormulaCount = nativeFormulaCount + 1
            Set formulaRange = candidateBookmark.Range.Duplicate
            If VTTryFormulaIdFromNativeBookmark( _
               candidateBookmark.Name, formulaId) Then
                If VTTryReadWordFormulaFormat( _
                   documentObject, formulaId, displayMode, numbered) Then
                    If displayMode = "block" Then
                        nativeDisplayCount = nativeDisplayCount + 1
                        Set paragraphRange = _
                            formulaRange.Paragraphs(1).Range.Duplicate
                        If CDbl(paragraphRange.ParagraphFormat.SpaceBefore) < _
                           minimumNativeSpaceBefore Then
                            minimumNativeSpaceBefore = CDbl( _
                                paragraphRange.ParagraphFormat.SpaceBefore)
                        End If
                        If CDbl(paragraphRange.ParagraphFormat.SpaceBefore) > _
                           maximumNativeSpaceBefore Then
                            maximumNativeSpaceBefore = CDbl( _
                                paragraphRange.ParagraphFormat.SpaceBefore)
                        End If
                        If CDbl(paragraphRange.ParagraphFormat.SpaceAfter) < _
                           minimumNativeSpaceAfter Then
                            minimumNativeSpaceAfter = CDbl( _
                                paragraphRange.ParagraphFormat.SpaceAfter)
                        End If
                        If CDbl(paragraphRange.ParagraphFormat.SpaceAfter) > _
                           maximumNativeSpaceAfter Then
                            maximumNativeSpaceAfter = CDbl( _
                                paragraphRange.ParagraphFormat.SpaceAfter)
                        End If
                    End If
                End If
            End If
            If formulaRange.OMaths.Count <> 1 Then
                invalidNativeRangeCount = invalidNativeRangeCount + 1
            Else
                formulaText = formulaRange.Text
                relationCount = 0
                minimumRelationPosition = 1000000#
                maximumRelationPosition = -1000000#
                For relationIndex = 1 To Len(formulaText)
                    If Mid$(formulaText, relationIndex, 1) = "=" Then
                        Set relationRange = documentObject.Range( _
                            Start:=formulaRange.Start + relationIndex - 1, _
                            End:=formulaRange.Start + relationIndex)
                        relationPosition = CDbl(relationRange.Information( _
                            wdHorizontalPositionRelativeToPage))
                        If relationPosition >= 0# Then
                            relationCount = relationCount + 1
                            If relationPosition < minimumRelationPosition Then
                                minimumRelationPosition = relationPosition
                            End If
                            If relationPosition > maximumRelationPosition Then
                                maximumRelationPosition = relationPosition
                            End If
                        End If
                    End If
                Next relationIndex
                If relationCount >= 2 Then
                    alignedFormulaCount = alignedFormulaCount + 1
                    If maximumRelationPosition - minimumRelationPosition > _
                       maximumAlignmentError Then
                        maximumAlignmentError = _
                            maximumRelationPosition - minimumRelationPosition
                    End If
                End If
            End If
        End If
    Next candidateBookmark

    For Each candidateShape In documentObject.InlineShapes
        If VTIsVisualTeXInlineShape(candidateShape) Then
            imageFormulaCount = imageFormulaCount + 1
            Set imageMacroField = _
                VTVisualTeXImageMacroButtonFieldForShape(candidateShape)
            If Not imageMacroField Is Nothing Then
                imageMacroButtonCount = imageMacroButtonCount + 1
                If Not VTMacroButtonResultHasOnlyImage(imageMacroField) Then
                    invalidImageMacroButtonCount = _
                        invalidImageMacroButtonCount + 1
                End If
            End If
            If VTTryParseFormulaReference( _
               candidateShape.Title, formulaId, displayMode, numbered) Then
                If Len(imageFormulaIds) > 0 Then
                    imageFormulaIds = imageFormulaIds & ","
                End If
                imageFormulaIds = imageFormulaIds & formulaId
                If displayMode = "block" Then
                    imageDisplayCount = imageDisplayCount + 1
                    Set paragraphRange = _
                        candidateShape.Range.Paragraphs(1).Range.Duplicate
                    If Abs(CDbl(paragraphRange.ParagraphFormat.SpaceBefore)) > _
                       maximumImageSpaceBefore Then
                        maximumImageSpaceBefore = Abs(CDbl( _
                            paragraphRange.ParagraphFormat.SpaceBefore))
                    End If
                    If Abs(CDbl(paragraphRange.ParagraphFormat.SpaceAfter)) > _
                       maximumImageSpaceAfter Then
                        maximumImageSpaceAfter = Abs(CDbl( _
                            paragraphRange.ParagraphFormat.SpaceAfter))
                    End If
                End If
            End If
        End If
    Next candidateShape

    If nativeDisplayCount = 0 Then
        minimumNativeSpaceBefore = 0#
        minimumNativeSpaceAfter = 0#
    End If

    report = "PASS" & vbLf
    report = report & "revision=" & VT_WORD_SOURCE_REVISION & vbLf
    report = report & "documentMathCount=" & _
        CStr(documentObject.OMaths.Count) & vbLf
    report = report & "nativeFormulaCount=" & _
        CStr(nativeFormulaCount) & vbLf
    report = report & "nativeDisplayCount=" & _
        CStr(nativeDisplayCount) & vbLf
    report = report & "invalidNativeRangeCount=" & _
        CStr(invalidNativeRangeCount) & vbLf
    report = report & "emptyMathCount=" & CStr(emptyMathCount) & vbLf
    report = report & "alignedFormulaCount=" & _
        CStr(alignedFormulaCount) & vbLf
    report = report & "maximumAlignmentError=" & _
        CStr(maximumAlignmentError) & vbLf
    report = report & "minimumNativeSpaceBefore=" & _
        CStr(minimumNativeSpaceBefore) & vbLf
    report = report & "maximumNativeSpaceBefore=" & _
        CStr(maximumNativeSpaceBefore) & vbLf
    report = report & "minimumNativeSpaceAfter=" & _
        CStr(minimumNativeSpaceAfter) & vbLf
    report = report & "maximumNativeSpaceAfter=" & _
        CStr(maximumNativeSpaceAfter) & vbLf
    report = report & "imageFormulaCount=" & _
        CStr(imageFormulaCount) & vbLf
    report = report & "imageDisplayCount=" & _
        CStr(imageDisplayCount) & vbLf
    report = report & "imageMacroButtonCount=" & _
        CStr(imageMacroButtonCount) & vbLf
    report = report & "invalidImageMacroButtonCount=" & _
        CStr(invalidImageMacroButtonCount) & vbLf
    report = report & "imageFormulaIds=" & imageFormulaIds & vbLf
    report = report & "maximumImageSpaceBefore=" & _
        CStr(maximumImageSpaceBefore) & vbLf
    report = report & "maximumImageSpaceAfter=" & _
        CStr(maximumImageSpaceAfter) & vbLf
    VTWriteTextAtomic statusPath, report
    Exit Sub

Failed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic statusPath, _
        "FAIL" & vbLf & _
        "revision=" & VT_WORD_SOURCE_REVISION & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    On Error GoTo 0
End Sub

Public Sub VisualTeX_ConfigureDocumentImportParagraphSpacingRegression()
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7592, "VisualTeX", _
            "The document-import spacing regression requires an active document."
    End If
    With ActiveDocument.Styles(wdStyleNormal).ParagraphFormat
        .SpaceBefore = 6!
        .SpaceAfter = 9!
    End With
End Sub

Private Function VTAppendRegressionParagraph( _
    ByVal documentObject As Document) As Range

    Dim insertionStart As Long
    Dim paragraphRange As Range

    If documentObject Is Nothing Then
        Err.Raise vbObjectError + 7548, "VisualTeX", _
            "The Word regression document is missing."
    End If
    documentObject.Content.InsertAfter vbCr
    insertionStart = documentObject.Content.End - 1
    Set paragraphRange = documentObject.Range( _
        Start:=insertionStart, End:=insertionStart).Paragraphs(1).Range.Duplicate
    VTNormalizePlainWordParagraph paragraphRange
    Set VTAppendRegressionParagraph = documentObject.Range( _
        Start:=insertionStart, End:=insertionStart)
End Function

Private Sub VTRegressionAssertImageNumberVerticalAlignment( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal expectedOrdinal As Long, _
    ByVal stageName As String, _
    ByRef numberPosition As Long)

    Dim formulaRange As Range
    Dim paragraphRange As Range
    Dim helperParagraph As Range
    Dim formulaShape As InlineShape
    Dim formulaContainerRange As Range
    Dim sequenceField As Field
    Dim visibleNumberField As Field
    Dim numberRange As Range
    Dim prefixRange As Range
    Dim separatorRange As Range
    Dim suffixRange As Range
    Dim expectedPosition As Long
    Dim centerDeltaPoints As Double
    Dim numberSize As Double
    Dim numberFontName As String
    Dim expectedNumberText As String
    Dim assertionStage As String
    Dim assertionErrorNumber As Long
    Dim assertionErrorDescription As String

    On Error GoTo AssertionFailed
    assertionStage = "validate-target"
    If documentObject Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Or _
       expectedOrdinal < 1 Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & ": the image-number alignment target is invalid."
    End If

    assertionStage = "resolve-formula"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & ": the regression formula Range is missing."
    End If
    If formulaRange.InlineShapes.Count <> 1 Or _
       formulaRange.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & ": the regression formula is not one image."
    End If
    Set formulaShape = formulaRange.InlineShapes(1)
    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate

    assertionStage = "resolve-number"
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    Set visibleNumberField = VTImageEquationReferenceField( _
        formulaRange, formulaId)
    If Not visibleNumberField Is Nothing Then
        Set numberRange = VTImageEquationNumberRange( _
            formulaRange, visibleNumberField)
    End If
    If sequenceField Is Nothing Or visibleNumberField Is Nothing Or _
       numberRange Is Nothing Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & ": the external SEQ or visible REF identity is missing."
    End If
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If helperParagraph.Start <> paragraphRange.End Or _
       Not VTHelperParagraphOwnsNativeEquationSequence( _
           helperParagraph) Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & ": the external SEQ helper is not immediately after the image row."
    End If

    assertionStage = "verify-structure"
    expectedNumberText = VTEquationNumberTextForFormula( _
        documentObject, formulaId)
    Set prefixRange = documentObject.Range( _
        Start:=paragraphRange.Start, End:=formulaContainerRange.Start)
    Set separatorRange = documentObject.Range( _
        Start:=formulaContainerRange.End, End:=numberRange.Start)
    Set suffixRange = documentObject.Range( _
        Start:=numberRange.End, End:=paragraphRange.End)
    If Len(expectedNumberText) = 0 Or _
       Not VTParagraphHasSingleVisualTeXImageMacroButton( _
           paragraphRange) Or _
       prefixRange.Text <> vbTab Or _
       separatorRange.Text <> vbTab Or _
       suffixRange.Text <> vbCr Or _
       numberRange.Text <> "(" & expectedNumberText & ")" Or _
       VTEquationSequenceResultText(sequenceField) <> _
           CStr(expectedOrdinal) Or _
       VTComparableEquationNumberText( _
           VTEquationSequenceResultText(visibleNumberField)) <> _
       VTComparableEquationNumberText(expectedNumberText) Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & _
            ": the layout is not <TAB><image><TAB>(REF)<CR> + adjacent SEQ helper" & _
            " [imageHost=" & CStr( _
                VTParagraphHasSingleVisualTeXImageMacroButton(paragraphRange)) & _
            "; prefix=" & Replace$(Replace$(prefixRange.Text, vbTab, "<TAB>"), vbCr, "<CR>") & _
            "; separator=" & Replace$(Replace$(separatorRange.Text, vbTab, "<TAB>"), vbCr, "<CR>") & _
            "; suffix=" & Replace$(Replace$(suffixRange.Text, vbTab, "<TAB>"), vbCr, "<CR>") & _
            "; number=" & numberRange.Text & _
            "; helperSEQ=" & VTEquationSequenceResultText(sequenceField) & _
            "; visibleREF=" & VTEquationSequenceResultText(visibleNumberField) & "]."
    End If

    assertionStage = "verify-tabs"
    If paragraphRange.ParagraphFormat.TabStops.Count <> 2 Or _
       paragraphRange.ParagraphFormat.TabStops(1).Alignment <> _
           wdAlignTabCenter Or _
       paragraphRange.ParagraphFormat.TabStops(2).Alignment <> _
           wdAlignTabRight Or _
       Abs(CDbl(paragraphRange.ParagraphFormat.TabStops(1).Position) - _
           207.65#) > 0.2# Or _
       Abs(CDbl(paragraphRange.ParagraphFormat.TabStops(2).Position) - _
           414.3#) > 0.2# Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & _
            ": the 207.65/414.30 point tab stops changed."
    End If

    assertionStage = "verify-paragraph-geometry"
    If Abs(CDbl(paragraphRange.ParagraphFormat.LeftIndent)) > 0.05# Or _
       Abs(CDbl(paragraphRange.ParagraphFormat.RightIndent)) > 0.05# Or _
       Abs(CDbl(paragraphRange.ParagraphFormat.FirstLineIndent)) > 0.05# Or _
       Abs(CDbl(paragraphRange.ParagraphFormat.LineSpacing) - 12#) > 0.1# Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & _
            ": the zero-indent, 12-point line geometry changed."
    End If

    assertionStage = "verify-image-geometry"
    If Abs(CDbl(formulaShape.Width) - 54.25#) > 0.1# Or _
       Abs(CDbl(formulaShape.Height) - 46.1#) > 0.1# Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & _
            ": the 54.25 x 46.10 point image fixture changed."
    End If

    assertionStage = "verify-number-font"
    numberSize = CDbl(numberRange.Font.Size)
    numberFontName = CStr(numberRange.Font.NameAscii)
    If numberSize <= 0# Or numberSize > 512# Or _
       Abs(numberSize - 10#) > 0.1# Or _
       numberRange.Font.Position <> 0 Or _
       StrComp(numberFontName, VT_WORD_EQUATION_NUMBER_FONT_NAME, _
           vbTextCompare) <> 0 Or _
       StrComp(CStr(numberRange.Font.NameOther), _
           VT_WORD_EQUATION_NUMBER_FONT_NAME, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & _
            ": the visible REF number does not use uniform 10-point" & _
            " Cambria Math at Position 0 [size=" & CStr(numberSize) & _
            "; ascii=" & numberFontName & _
            "; other=" & CStr(numberRange.Font.NameOther) & _
            "; position=" & CStr(numberRange.Font.Position) & "]."
    End If

    assertionStage = "verify-number-position"
    ' Word for Mac can throw runtime error 6 when Font.Position is read in the
    ' transient layout window immediately after a field refresh. Reapply the
    ' deterministic value and use the validated calculation as the source of
    ' truth instead of reading the unstable property back synchronously.
    expectedPosition = VTApplyNumberedImageFormulaVerticalAlignment( _
        formulaShape, numberRange, CSng(numberSize), _
        stageName & ": the measured image-object Position")
    If Not VTTryCalculateCurrentNumberedImagePaintedCenterDeltaPoints( _
       formulaShape, numberRange, centerDeltaPoints) Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & ": Word did not expose final ink-center geometry."
    End If
    If Abs(centerDeltaPoints) > 0.5001# Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & ": final ink centers exceed Position quantisation" & _
            " [deltaPoints=" & CStr(centerDeltaPoints) & "]."
    End If
    numberPosition = Abs(formulaShape.Range.Font.Position)
    Exit Sub

AssertionFailed:
    assertionErrorNumber = Err.Number
    assertionErrorDescription = Err.Description
    If assertionErrorNumber = 0 Then assertionErrorNumber = vbObjectError + 7568
    Err.Raise assertionErrorNumber, "VisualTeX regression", _
        stageName & "/" & assertionStage & ": " & _
        assertionErrorDescription
End Sub

Public Sub VisualTeX_RunWordImageNumberVerticalAlignmentRegression()
    Const formulaId As String = _
        "28282828-2828-4828-8828-282828282828"

    Dim sourceDocument As Document
    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim insertionRange As Range
    Dim paragraphRange As Range
    Dim helperParagraph As Range
    Dim sequenceField As Field
    Dim numberRange As Range
    Dim selectedNumberRange As Range
    Dim encodedMetadata As String
    Dim latexBase64 As String
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim positionBefore As Long
    Dim positionAt24 As Long
    Dim expectedPositionAt24 As Long
    Dim heightAt24 As Double
    Dim numberSizeAt24 As Double
    Dim expectedNumberSizeAt24 As Double
    Dim manualResizeHeight As Double
    Dim expectedManualResizePosition As Long
    Dim positionAfterSelection As Long
    Dim positionAfter As Long
    Dim ignoredReturnRepair As Boolean
    Dim report As String
    Dim failureReport As String

    On Error GoTo RegressionFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-image-number-vertical-alignment-regression-result.txt"
    If Documents.Count > 0 Then Set sourceDocument = ActiveDocument

    regressionStage = "create-54p25-by-46p10-fixture"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    testDocument.ActiveWindow.View.Type = wdPrintView
    With testDocument.PageSetup
        .Orientation = wdOrientPortrait
        .PageWidth = CentimetersToPoints(21#)
        .PageHeight = CentimetersToPoints(29.7)
        .LeftMargin = 90!
        .RightMargin = 90!
        .TextColumns.SetCount NumColumns:=1
    End With
    testDocument.Styles(wdStyleNormal).Font.Size = 10!
    testDocument.Styles(wdStyleCaption).Font.Size = 10!
    testDocument.Repaginate
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    Set formulaShape = testDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    formulaShape.LockAspectRatio = msoFalse
    formulaShape.Width = 54.25!
    formulaShape.Height = 46.1!
    formulaShape.Range.Font.Position = 0
    encodedMetadata = VT_METADATA_PREFIX & "e30"
    latexBase64 = "eF8y"
    formulaShape.AlternativeText = encodedMetadata
    formulaShape.Title = VTFormulaReference(formulaId, "block", True)
    VTSetWordLatexPayload testDocument, formulaId, latexBase64
    VTSetWordMetadataPayload _
        testDocument, formulaId, encodedMetadata
    VTSetWordFormulaFormat testDocument, formulaId, "block", True
    ' The 10 pt fixture is 54.25 x 46.10 pt. Store its equivalent 14 pt
    ' reference geometry and a -21 pt SVG baseline, so the expected number
    ' raises are 15 pt at 10 pt and 36 pt at 24 pt.
    VTSetWordImageScaleState _
        testDocument, formulaId, 10#, 75.95#, 64.54#, -21#, 10#
    Set paragraphRange = VTInsertEquationNumber( _
        formulaShape, formulaId, "vertical alignment fixture")

    regressionStage = "calibrate-before-field-refresh"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    With paragraphRange.ParagraphFormat
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 12!
    End With
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, formulaId)
    Set numberRange = VTStaticImageEquationNumberRange( _
        formulaRange, formulaId)
    If sequenceField Is Nothing Or numberRange Is Nothing Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The image-number regression external SEQ/visible REF identity is missing."
    End If
    With numberRange.Font
        .Size = 10!
        .Position = 0
    End With
    formulaRange.InlineShapes(1).Range.Font.Position = 0
    VTSetEquationNumberBookmarkExact _
        testDocument, formulaId, numberRange
    positionBefore = VTCalculateStaticImageEquationNumberPosition( _
        formulaRange.InlineShapes(1), numberRange)
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    VTAssertNumberedEquationLayout _
        formulaRange, 46.1!, formulaId, _
        "vertical alignment fixture", _
        "image-number baseline regression before field refresh"
    VTVerifyParagraphEquationNumberIntegrity _
        formulaRange, formulaId, 1
    VTRegressionAssertImageNumberVerticalAlignment _
        testDocument, formulaId, 1, _
        "before field refresh", positionBefore

    regressionStage = "numbered-image-enter-creates-normal-line"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    If paragraphRange.ParagraphFormat.Alignment <> wdAlignParagraphJustify Or _
       VTCustomTabStopCount(paragraphRange) <> 2 Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The numbered image formula does not use the expected justified center/right-tab host."
    End If
    Set insertionRange = testDocument.Range( _
        Start:=paragraphRange.End - 1, End:=paragraphRange.End - 1)
    insertionRange.Select
    Selection.TypeParagraph
    DoEvents
    ' Programmatic TypeParagraph does not reliably raise WindowSelectionChange on
    ' Word for Mac. Invoke the exact same O(1) Return repair here so this VBA
    ' regression validates the structural transaction itself; a separate physical
    ' keyboard regression verifies the real SelectionChange event path.
    ignoredReturnRepair = _
        VTRepairNumberedImageReturnAfterFormula(Selection.Range)
    DoEvents
    Set paragraphRange = Selection.Range.Paragraphs(1).Range.Duplicate
    If paragraphRange.Information(wdWithInTable) Or _
       paragraphRange.Fields.Count <> 0 Or _
       paragraphRange.InlineShapes.Count <> 0 Or _
       paragraphRange.OMaths.Count <> 0 Or _
       VTCustomTabStopCount(paragraphRange) <> 0 Or _
       paragraphRange.Font.Position <> 0 Or _
       VTWordRangeHasMeaningfulText(paragraphRange) Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "Pressing Return after a numbered image formula did not create one plain Normal paragraph" & _
            " [range=" & CStr(paragraphRange.Start) & "-" & CStr(paragraphRange.End) & _
            "; style=" & CStr(paragraphRange.Style) & _
            "; alignment=" & CStr(paragraphRange.ParagraphFormat.Alignment) & _
            "; customTabs=" & CStr(VTCustomTabStopCount(paragraphRange)) & _
            "; fields=" & CStr(paragraphRange.Fields.Count) & _
            "; images=" & CStr(paragraphRange.InlineShapes.Count) & _
            "; omaths=" & CStr(paragraphRange.OMaths.Count) & _
            "; position=" & CStr(paragraphRange.Font.Position) & _
            "; text=" & Replace$(Replace$(paragraphRange.Text, vbCr, "<CR>"), _
                vbTab, "<TAB>") & "]."
    End If
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, formulaId)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The plain typing paragraph lost the image Equation SEQ helper."
    End If
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If VTEquationSequenceResultText(sequenceField) <> "1" Or _
       helperParagraph.Start <> paragraphRange.End Or _
       paragraphRange.Start >= helperParagraph.Start Or _
       Selection.Range.Paragraphs(1).Range.Start <> helperParagraph.End Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "Return did not preserve formula -> helper -> typing-line order" & _
            " [formula=" & CStr(paragraphRange.Start) & "-" & _
                CStr(paragraphRange.End) & _
            "; helper=" & CStr(helperParagraph.Start) & "-" & _
                CStr(helperParagraph.End) & _
            "; typing=" & CStr(Selection.Range.Paragraphs(1).Range.Start) & "]."
    End If
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)

    regressionStage = "resize-numbered-image-to-24"
    Set formulaShape = formulaRange.InlineShapes(1)
    VTApplyWordImageFormulaFontSize formulaShape, 24#
    DoEvents
    testDocument.Repaginate
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    Set numberRange = VTStaticImageEquationNumberRange( _
        formulaRange, formulaId)
    If numberRange Is Nothing Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The 24-point image resize lost its visible number."
    End If
    heightAt24 = CDbl(formulaRange.InlineShapes(1).Height)
    numberSizeAt24 = CDbl(numberRange.Font.Size)
    expectedNumberSizeAt24 = _
        CDbl(VTVisibleEquationNumberFontSize(testDocument))
    If Abs(numberSizeAt24 - expectedNumberSizeAt24) > 0.1 Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The 24-point image resize changed the document Equation number size" & _
            " [actual=" & CStr(numberSizeAt24) & _
            "; expected=" & CStr(expectedNumberSizeAt24) & "]."
    End If
    ' The complete layout assertion verifies Word's applied number Position,
    ' centre tab and right tab. Do not synchronously read Font.Position again.
    VTAssertNumberedEquationLayout _
        formulaRange, heightAt24, formulaId, _
        "vertical alignment fixture", _
        "image-number resize regression at 24 pt"
    expectedPositionAt24 = Abs( _
        formulaRange.InlineShapes(1).Range.Font.Position)
    positionAt24 = expectedPositionAt24

    regressionStage = "resize-numbered-image-back-to-10"
    Set formulaShape = formulaRange.InlineShapes(1)
    VTApplyWordImageFormulaFontSize formulaShape, 10#
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    VTRegressionAssertImageNumberVerticalAlignment _
        testDocument, formulaId, 1, _
        "after resizing 24 pt back to 10 pt", positionBefore

    regressionStage = "direct-resize-numbered-image"
    Set formulaShape = formulaRange.InlineShapes(1)
    formulaShape.LockAspectRatio = msoFalse
    formulaShape.Width = 65.1!
    formulaShape.Height = 55.32!
    formulaShape.LockAspectRatio = msoTrue
    VTSynchronizeWordImageFormulaShape formulaShape
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    manualResizeHeight = CDbl(formulaRange.InlineShapes(1).Height)
    Set numberRange = VTStaticImageEquationNumberRange( _
        formulaRange, formulaId)
    If numberRange Is Nothing Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "Direct image resizing lost the visible Equation number."
    End If
    VTAssertNumberedEquationLayout _
        formulaRange, manualResizeHeight, formulaId, _
        "vertical alignment fixture", _
        "image-number direct-resize baseline regression"
    expectedManualResizePosition = Abs( _
        formulaRange.InlineShapes(1).Range.Font.Position)

    regressionStage = "restore-direct-resized-image"
    Set formulaShape = formulaRange.InlineShapes(1)
    formulaShape.LockAspectRatio = msoFalse
    formulaShape.Width = 54.25!
    formulaShape.Height = 46.1!
    formulaShape.LockAspectRatio = msoTrue
    VTSynchronizeWordImageFormulaShape formulaShape
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    VTRegressionAssertImageNumberVerticalAlignment _
        testDocument, formulaId, 1, _
        "after restoring direct image resize", positionBefore

    regressionStage = "select-static-visible-number"
    Set numberRange = VTStaticImageEquationNumberRange( _
        formulaRange, formulaId)
    If numberRange Is Nothing Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The static image Equation number disappeared before selection."
    End If
    Set selectedNumberRange = numberRange.Duplicate
    selectedNumberRange.Start = selectedNumberRange.Start + 1
    selectedNumberRange.End = selectedNumberRange.End - 1
    selectedNumberRange.Select
    DoEvents
    testDocument.Repaginate
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    Set numberRange = VTStaticImageEquationNumberRange( _
        formulaRange, formulaId)
    If numberRange Is Nothing Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "Selecting the static image Equation number damaged its range."
    End If
    VTRegressionAssertImageNumberVerticalAlignment _
        testDocument, formulaId, 1, _
        "after selecting visible number result", _
        positionAfterSelection
    If positionAfterSelection <> positionBefore Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "Selecting the image Equation number changed its baseline" & _
            " [before=" & CStr(positionBefore) & _
            "; after=" & CStr(positionAfterSelection) & "]."
    End If

    regressionStage = "refresh-fields"
    testDocument.Fields.Update
    regressionStage = "reconcile-image-number-after-field-refresh"
    VTReconcileEquationNumbers testDocument
    regressionStage = "repaginate-image-number-after-field-refresh"
    testDocument.Repaginate
    regressionStage = "resolve-image-formula-after-field-refresh"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    regressionStage = "verify-static-image-number-after-field-refresh"
    Set numberRange = VTStaticImageEquationNumberRange( _
        formulaRange, formulaId)
    If numberRange Is Nothing Or _
       Not VTParagraphHasSingleVisualTeXImageMacroButton( _
           formulaRange.Paragraphs(1).Range) Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The refreshed image Equation visible REF is incomplete."
    End If
    regressionStage = "assert-image-layout-after-field-refresh"
    VTAssertNumberedEquationLayout _
        formulaRange, 46.1!, formulaId, _
        "vertical alignment fixture", _
        "image-number baseline regression after field refresh"
    regressionStage = "verify-image-integrity-after-field-refresh"
    VTVerifyParagraphEquationNumberIntegrity _
        formulaRange, formulaId, 1
    regressionStage = "record-image-vertical-after-field-refresh"
    ' The complete layout and integrity assertions immediately above already
    ' verified the original 54.25 x 46.10 fixture, 10 pt number, centre tab,
    ' right tab and deterministic vertical correction. Record the previously
    ' validated target instead of reading any transient Word layout property.
    positionAfter = positionBefore

    ' Record the successful result before attempting any Word UI/document
    ' cleanup. Word for Mac can raise runtime error 6 while closing the
    ' repaginated temporary document even though every formula assertion has
    ' already passed; cleanup must never overwrite a valid regression result.
    regressionStage = "build-success-report"
    report = "PASS" & vbLf
    report = report & "revision=" & VT_WORD_SOURCE_REVISION & vbLf
    report = report & _
        "structure=<TAB><image><TAB>(REF)<CR>+adjacent-SEQ-helper" & vbLf
    report = report & "image=54.25x46.10" & vbLf
    report = report & "fontSize=10" & vbLf
    report = report & "lineSpacing=12" & vbLf
    report = report & "centerTab=207.65" & vbLf
    report = report & "rightTab=414.30" & vbLf
    report = report & "positionBefore=" & CStr(positionBefore) & vbLf
    report = report & "heightAt24=" & CStr(heightAt24) & vbLf
    report = report & "positionAt24=" & CStr(positionAt24) & vbLf
    report = report & _
        "expectedPositionAt24=" & CStr(expectedPositionAt24) & vbLf
    report = report & _
        "manualResizeHeight=" & CStr(manualResizeHeight) & vbLf
    report = report & _
        "expectedManualResizePosition=" & _
            CStr(expectedManualResizePosition) & vbLf
    report = report & "positionAfterSelection=" & _
        CStr(positionAfterSelection) & vbLf
    report = report & "positionAfter=" & CStr(positionAfter) & vbLf
    report = report & _
        "alignmentFormula=round(CambriaMathGlyphCenter-PngInkCenterFromBaseline)" & vbLf
    report = report & _
        "numberFont=" & VT_WORD_EQUATION_NUMBER_FONT_NAME & vbLf
    report = report & _
        "initialVisualCenterShift=" & CStr(positionBefore) & vbLf
    regressionStage = "write-success-report"
    VTWriteTextAtomic resultPath, report

    regressionStage = "cleanup-successful-regression"
    On Error Resume Next
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set testDocument = Nothing
    End If
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    Err.Clear
    On Error GoTo 0
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    If regressionErrorNumber = 0 Then
        regressionErrorNumber = vbObjectError + 7568
    End If
    If Len(regressionErrorDescription) = 0 Then
        regressionErrorDescription = _
            "The Word regression stopped without exposing an error description."
    End If
    failureReport = "FAIL" & vbLf
    failureReport = failureReport & _
        "revision=" & VT_WORD_SOURCE_REVISION & vbLf
    failureReport = failureReport & _
        "stage=" & regressionStage & vbLf
    failureReport = failureReport & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf
    failureReport = failureReport & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    On Error Resume Next
    VTWriteTextAtomic resultPath, failureReport
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    On Error GoTo 0
    Err.Raise regressionErrorNumber, "VisualTeX regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordSingleParagraphNumberRegression()
    Const formulaId As String = _
        "11111111-1111-4111-8111-111111111111"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim paragraphRange As Range
    Dim insertionRange As Range
    Dim referenceRange As Range
    Dim encodedMetadata As String
    Dim latexBase64 As String
    Dim ommlBase64 As String
    Dim fixtureRoot As String
    Dim nativeDocumentPath As String
    Dim referenceResult As String
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim tabCount As Long

    On Error GoTo RegressionFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    nativeDocumentPath = VTApplicationSupportRoot() & _
        "/NativeDocuments/" & formulaId & ".docx"
    resultPath = fixtureRoot & _
        "/word-single-paragraph-number-regression-result.txt"
    encodedMetadata = VT_METADATA_PREFIX & "e30"
    latexBase64 = "eF8x"
    If Not VTPathFileExists(nativeDocumentPath) Then
        Err.Raise vbObjectError + 7562, "VisualTeX", _
            "The single-paragraph conversion fixture DOCX is missing."
    End If
    ommlBase64 = VTReadText( _
        fixtureRoot & "/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    If Len(ommlBase64) = 0 Then
        Err.Raise vbObjectError + 7562, "VisualTeX", _
            "The single-paragraph conversion OMML fixture is missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."

    regressionStage = "blank-first-line-create"
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    Set formulaShape = testDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    formulaShape.Width = 120!
    formulaShape.Height = 36!
    formulaShape.AlternativeText = encodedMetadata
    formulaShape.Title = VTFormulaReference(formulaId, "block", True)
    VTSetWordLatexPayload testDocument, formulaId, latexBase64
    VTSetWordOmmlPayload testDocument, formulaId, ommlBase64
    VTSetWordMetadataPayload testDocument, formulaId, encodedMetadata
    VTSetWordFormulaFormat testDocument, formulaId, "block", True
    Set paragraphRange = VTInsertEquationNumber( _
        formulaShape, formulaId, _
        VTEquationCrossReferenceText(latexBase64))

    regressionStage = "blank-first-line-layout"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7562, "VisualTeX", _
            "A first-line numbered formula lost its formula Range."
    End If
    If formulaRange.Information(wdWithInTable) Or _
       testDocument.Tables.Count <> 0 Then
        Err.Raise vbObjectError + 7562, "VisualTeX", _
            "A first-line numbered formula created a table."
    End If
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    tabCount = Len(paragraphRange.Text) - _
        Len(Replace$(paragraphRange.Text, vbTab, ""))
    If paragraphRange.Paragraphs.Count <> 1 Or tabCount <> 2 Or _
       paragraphRange.ParagraphFormat.KeepTogether Or _
       paragraphRange.ParagraphFormat.PageBreakBefore Then
        Err.Raise vbObjectError + 7562, "VisualTeX", _
            "A first-line numbered image is not one clean paragraph with two tabs" & _
            " [paragraphs=" & CStr(paragraphRange.Paragraphs.Count) & _
            "; tabs=" & CStr(tabCount) & _
            "; keepTogether=" & _
                CStr(paragraphRange.ParagraphFormat.KeepTogether) & "]."
    End If
    VTAssertNumberedEquationLayout _
        formulaRange, 36!, formulaId, _
        VTEquationCrossReferenceText(latexBase64), _
        "blank first-line single-paragraph formula"
    VTVerifyParagraphEquationNumberIntegrity _
        formulaRange, formulaId, 1

    regressionStage = "blank-first-line-cross-reference"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set referenceRange = VTInsertEquationNumberReferenceAtRange( _
        insertionRange, 1)
    If referenceRange.Text <> "(1)" Or _
       referenceRange.Fields.Count <> 1 Then
        Err.Raise vbObjectError + 7562, "VisualTeX", _
            "The first-line formula did not expose a live native reference (1)."
    End If
    VTAssertBodyEquationReferenceVisible _
        testDocument, referenceRange.Fields(1), _
        "blank first-line Equation reference"
    VTVerifyNumberedFormulaIntegrity testDocument, formulaId, 1

    regressionStage = "blank-first-line-convert-to-omml"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7562, "VisualTeX", _
            "The numbered image formula disappeared before native conversion."
    End If
    If formulaRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7562, "VisualTeX", _
            "The numbered image conversion target is ambiguous."
    End If
    Set formulaShape = formulaRange.InlineShapes(1)
    VTWordConvertInlineShapeToNativeEquation formulaShape

    regressionStage = "converted-omml-layout"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7562, "VisualTeX", _
            "Image-to-OMML conversion lost the numbered formula."
    End If
    If formulaRange.OMaths.Count <> 1 Or _
       formulaRange.OMaths(1).Type <> wdOMathDisplay Or _
       formulaRange.InlineShapes.Count <> 0 Or _
       formulaRange.Information(wdWithInTable) Or _
       testDocument.Tables.Count <> 0 Then
        Err.Raise vbObjectError + 7562, "VisualTeX", _
            "Image-to-OMML conversion did not preserve one table-free numbered formula."
    End If
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    tabCount = Len(paragraphRange.Text) - _
        Len(Replace$(paragraphRange.Text, vbTab, ""))
    If paragraphRange.Paragraphs.Count <> 1 Or tabCount <> 0 Or _
       paragraphRange.ParagraphFormat.KeepTogether Or _
       paragraphRange.ParagraphFormat.PageBreakBefore Then
        Err.Raise vbObjectError + 7562, "VisualTeX", _
            "Converted display OMML is not one clean Equation-array paragraph" & _
            " [paragraphs=" & CStr(paragraphRange.Paragraphs.Count) & _
            "; tabs=" & CStr(tabCount) & _
            "; keepTogether=" & _
                CStr(paragraphRange.ParagraphFormat.KeepTogether) & "]."
    End If
    VTAssertNumberedEquationLayout _
        formulaRange, 36!, formulaId, _
        VTEquationCrossReferenceText(latexBase64), _
        "converted single-paragraph OMML formula"
    VTVerifyParagraphEquationNumberIntegrity _
        formulaRange, formulaId, 1

    regressionStage = "converted-omml-refresh-and-reference"
    VTReconcileEquationNumbers testDocument
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    VTAssertNumberedEquationLayout _
        formulaRange, 36!, formulaId, _
        VTEquationCrossReferenceText(latexBase64), _
        "refreshed converted single-paragraph OMML formula"
    referenceResult = VTBodyReferenceResultForTarget( _
        testDocument, VTEquationSequenceNumberBookmarkName(formulaId))
    If referenceResult <> "1" Then
        Err.Raise vbObjectError + 7562, "VisualTeX", _
            "The converted OMML formula lost its dynamic reference" & _
            " [result=" & referenceResult & "]."
    End If

    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    VTWriteTextAtomic resultPath, _
        "PASS" & vbLf & _
        "tables=0" & vbLf & _
        "paragraphs=1" & vbLf & _
        "imageTabs=2" & vbLf & _
        "reference=(1)" & vbLf & _
        "convertedOmml=1" & vbLf & _
        "convertedTables=0" & vbLf & _
        "convertedTabs=0" & vbLf & _
        "convertedDisplay=PASS" & vbLf & _
        "convertedVisualCenter=PASS" & vbLf & _
        "convertedReference=(1)" & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    ' Preserve the failed document for popup/process/range inspection.
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word single-paragraph regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordUserWorkflowRegression()
    Const imageFormulaId As String = _
        "88888888-8888-4888-8888-888888888888"
    Const nativeFormulaId As String = _
        "99999999-9999-4999-8999-999999999999"
    Const fixtureFormulaId As String = _
        "11111111-1111-4111-8111-111111111111"

    Dim testDocument As Document
    Dim inlineFormula As InlineShape
    Dim displayFormula As InlineShape
    Dim nativeEquationRange As Range
    Dim insertionRange As Range
    Dim typedRange As Range
    Dim orphanRange As Range
    Dim imageTable As Table
    Dim nativeTable As Table
    Dim imageParagraph As Range
    Dim imageNumberRange As Range
    Dim fixtureRoot As String
    Dim nativeDocumentPath As String
    Dim ommlBase64 As String
    Dim imageMetadata As String
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim textStart As Long
    Dim normalSize As Single
    Dim numberCreated As Boolean

    On Error GoTo RegressionFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    nativeDocumentPath = _
        VTApplicationSupportRoot() & "/NativeDocuments/" & _
        fixtureFormulaId & ".docx"
    ommlBase64 = VTReadText( _
        fixtureRoot & "/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    resultPath = _
        fixtureRoot & "/word-user-workflow-regression-result.txt"
    If Not VTPathFileExists(nativeDocumentPath) Then
        Err.Raise vbObjectError + 7552, "VisualTeX", _
            "The workflow native DOCX fixture is missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    normalSize = VTVisibleEquationNumberFontSize(testDocument)

    regressionStage = "insert-inline-image"
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    Set inlineFormula = testDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    inlineFormula.Width = 70!
    inlineFormula.Height = 16!

    regressionStage = "prepare-image-display-after-inline"
    Set insertionRange = inlineFormula.Range.Duplicate
    insertionRange.Collapse wdCollapseEnd
    Set insertionRange = VTPrepareWordCreateInsertionRange( _
        insertionRange, "block")
    Set displayFormula = testDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    displayFormula.Width = 120!
    displayFormula.Height = 36!
    imageMetadata = VT_METADATA_PREFIX & "e30"
    displayFormula.AlternativeText = imageMetadata
    displayFormula.Title = _
        VTFormulaReference(imageFormulaId, "block", True)
    VTSetWordLatexPayload testDocument, imageFormulaId, "eF8x"
    VTSetWordMetadataPayload testDocument, imageFormulaId, imageMetadata
    If testDocument.InlineShapes.Count <> 2 Or _
       inlineFormula.Range.Paragraphs(1).Range.Start = _
           displayFormula.Range.Paragraphs(1).Range.Start Then
        Err.Raise vbObjectError + 7552, "VisualTeX", _
            "Creating a display formula after an inline formula did not" & _
            " preserve two independent paragraphs."
    End If

    regressionStage = "number-image-display-visible"
    VTSetWordFormulaFormat _
        testDocument, imageFormulaId, "block", True
    Set insertionRange = VTInsertEquationNumber( _
        displayFormula, imageFormulaId, "workflow image formula")
    VTVerifyNumberedFormulaIntegrity _
        testDocument, imageFormulaId, 1

    regressionStage = "simulate-legacy-image-number-table"
    Set imageParagraph = displayFormula.Range.Paragraphs(1).Range.Duplicate
    Set imageTable = imageParagraph.ConvertToTable( _
        Separator:=wdSeparateByTabs, NumRows:=1, NumColumns:=3)
    If imageTable.Rows.Count <> 1 Or imageTable.Columns.Count <> 3 Or _
       imageTable.Cell(1, 2).Range.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7552, "VisualTeX", _
            "The workflow regression could not construct one legacy numbered image table."
    End If
    Set displayFormula = imageTable.Cell(1, 2).Range.InlineShapes(1)

    regressionStage = "migrate-legacy-image-number-table"
    numberCreated = False
    Set imageNumberRange = VTEnsureImageEquationNumber( _
        displayFormula, displayFormula.Height, imageFormulaId, _
        "workflow image formula", numberCreated)
    Set imageParagraph = displayFormula.Range.Paragraphs(1).Range.Duplicate
    If testDocument.Tables.Count <> 0 Or _
       imageParagraph.Information(wdWithInTable) Or _
       VTCustomTabStopCount(imageParagraph) <> 2 Or _
       imageParagraph.InlineShapes.Count <> 1 Or _
       imageNumberRange.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 7552, "VisualTeX", _
            "The legacy numbered image table did not converge to one two-tab paragraph."
    End If
    VTVerifyNumberedFormulaIntegrity _
        testDocument, imageFormulaId, 1

    regressionStage = "image-display-continuation"
    VTPlaceCaretAfterDisplayFormula displayFormula.Range, imageFormulaId
    textStart = Selection.Start
    Selection.TypeText Text:="workflow continuation"
    Set typedRange = testDocument.Range( _
        Start:=textStart, End:=Selection.Start)
    If typedRange.Text <> "workflow continuation" Or _
       typedRange.Information(wdWithInTable) Or _
       typedRange.Font.Hidden <> False Or _
       typedRange.Font.Color <> wdColorAutomatic Or _
       Abs(typedRange.Font.Size - normalSize) > 0.1 Or _
       typedRange.ParagraphFormat.LineSpacingRule <> wdLineSpaceSingle Or _
       typedRange.ParagraphFormat.Alignment <> wdAlignParagraphLeft Then
        Err.Raise vbObjectError + 7552, "VisualTeX", _
            "The line after an image display formula is not ordinary body text."
    End If

    regressionStage = "insert-native-numbered-after-image"
    Set insertionRange = VTPrepareWordCreateInsertionRange( _
        Selection.Range.Duplicate, "block")
    Set nativeEquationRange = VTInsertNativeEquationAtRange( _
        insertionRange, ommlBase64, nativeDocumentPath, _
        "inline", True, False)
    VTSetWordFormulaFormat _
        testDocument, nativeFormulaId, "block", True
    numberCreated = False
    Set insertionRange = VTEnsureNativeEquationNumber( _
        nativeEquationRange, 48#, nativeFormulaId, _
        "workflow native formula", numberCreated)
    If Not numberCreated Then
        Err.Raise vbObjectError + 7552, "VisualTeX", _
            "The second numbered OMML formula did not create its own number."
    End If

    regressionStage = "verify-both-visible-after-second-insertion"
    VTReconcileEquationNumbers testDocument
    VTVerifyNumberedFormulaIntegrity _
        testDocument, imageFormulaId, 1
    VTVerifyNumberedFormulaIntegrity _
        testDocument, nativeFormulaId, 2

    regressionStage = "native-display-continuation"
    Set nativeEquationRange = VTNumberedFormulaRangeForId( _
        testDocument, nativeFormulaId)
    VTPlaceCaretAfterDisplayFormula nativeEquationRange, nativeFormulaId
    textStart = Selection.Start
    Selection.TypeText Text:="native continuation"
    Set typedRange = testDocument.Range( _
        Start:=textStart, End:=Selection.Start)
    If typedRange.Text <> "native continuation" Or _
       typedRange.Information(wdWithInTable) Or _
       typedRange.Font.Hidden <> False Or _
       typedRange.Font.Color <> wdColorAutomatic Or _
       Abs(typedRange.Font.Size - normalSize) > 0.1 Or _
       typedRange.ParagraphFormat.LineSpacingRule <> wdLineSpaceSingle Or _
       typedRange.ParagraphFormat.Alignment <> wdAlignParagraphLeft Then
        Err.Raise vbObjectError + 7552, "VisualTeX", _
            "The line after an OMML display formula is not ordinary body text."
    End If

    regressionStage = "delete-native-numbered-display"
    Set nativeEquationRange = VTNumberedFormulaRangeForId( _
        testDocument, nativeFormulaId)
    Set orphanRange = nativeEquationRange.Paragraphs(1).Range.Duplicate
    nativeEquationRange.Delete
    VTPruneOrphanedEquationNumberScaffolds testDocument
    VTReconcileEquationNumbers testDocument
    VTVerifyNumberedFormulaIntegrity _
        testDocument, imageFormulaId, 1

    regressionStage = "delete-image-numbered-display"
    Set orphanRange = VTNumberedFormulaRangeForId( _
        testDocument, imageFormulaId)
    orphanRange.InlineShapes(1).Delete
    VTPruneOrphanedEquationNumberScaffolds testDocument
    If testDocument.Tables.Count <> 0 Or _
       testDocument.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7552, "VisualTeX", _
            "Deleting numbered displays left an orphan table or removed" & _
            " the earlier inline formula."
    End If
    textStart = Selection.Start
    Selection.TypeText Text:="after cleanup"
    Set typedRange = testDocument.Range( _
        Start:=textStart, End:=Selection.Start)
    If typedRange.Text <> "after cleanup" Or _
       typedRange.Information(wdWithInTable) Or _
       typedRange.Font.Hidden <> False Or _
       Abs(typedRange.Font.Size - normalSize) > 0.1 Or _
       typedRange.ParagraphFormat.LineSpacingRule <> wdLineSpaceSingle Then
        Err.Raise vbObjectError + 7552, "VisualTeX", _
            "Orphan cleanup did not restore an ordinary body-text caret."
    End If

    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    VTWriteTextAtomic resultPath, "PASS" & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    ' Preserve the failed document for popup/process/range inspection.
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word user workflow regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Private Function VTRegressionCreateNumberedImage( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal latexBase64 As String, _
    ByVal widthPoints As Single, _
    ByVal heightPoints As Single, _
    Optional ByRef phaseTimings As Variant, _
    Optional ByRef numberingPhaseTimings As Variant) As InlineShape

    Dim insertionRange As Range
    Dim numberRange As Range
    Dim formulaShape As InlineShape
    Dim encodedMetadata As String
    Dim phaseStartedAt As Single
    Dim appendSeconds As Double
    Dim pictureSeconds As Double
    Dim metadataSeconds As Double
    Dim numberingSeconds As Double
    Dim resolveSeconds As Double

    encodedMetadata = VT_METADATA_PREFIX & "e30"
    phaseStartedAt = Timer
    Set insertionRange = VTAppendRegressionParagraph(documentObject)
    appendSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)

    phaseStartedAt = Timer
    Set formulaShape = documentObject.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    formulaShape.Width = widthPoints
    formulaShape.Height = heightPoints
    formulaShape.AlternativeText = encodedMetadata
    formulaShape.Title = VTFormulaReference(formulaId, "block", True)
    pictureSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)

    phaseStartedAt = Timer
    VTSetWordLatexPayload documentObject, formulaId, latexBase64
    VTSetWordMetadataPayload documentObject, formulaId, encodedMetadata
    VTSetWordFormulaFormat documentObject, formulaId, "block", True
    metadataSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)

    phaseStartedAt = Timer
    Set numberRange = VTInsertEquationNumber( _
        formulaShape, formulaId, _
        VTEquationCrossReferenceText(latexBase64), False, _
        numberingPhaseTimings)
    numberingSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)

    phaseStartedAt = Timer
    Set numberRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    resolveSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    phaseTimings = Array( _
        appendSeconds, pictureSeconds, metadataSeconds, _
        numberingSeconds, resolveSeconds)
    If numberRange Is Nothing Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The regression image formula identity is missing."
    End If
    If numberRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The regression image formula did not keep its single-paragraph identity."
    End If
    Set VTRegressionCreateNumberedImage = _
        numberRange.InlineShapes(1)
End Function

Public Sub VisualTeX_SetupWordImagePhysicalReturnRegression()
    Const formulaId As String = _
        "29292929-2929-4929-8929-292929292929"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim paragraphRange As Range
    Dim caretRange As Range
    Dim resultPath As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo SetupFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-image-physical-return-regression-result.txt"
    If Documents.Count <> 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The physical Return regression requires Word to have no open documents."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.Activate
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Styles(wdStyleNormal).Font.Size = 10!
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_MODE_VARIABLE, _
        VT_WORD_NUMBERING_MODE_SEQUENCE
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_SEPARATOR_VARIABLE, "."

    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, formulaId, "eF8y", 96!, 28!)
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The physical Return fixture lost its numbered image formula."
    End If
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    Set caretRange = testDocument.Range( _
        Start:=paragraphRange.End - 1, End:=paragraphRange.End - 1)
    caretRange.Select
    VTWriteTextAtomic _
        resultPath, _
        "READY" & vbLf & _
        "formula=" & CStr(paragraphRange.Start) & "-" & _
            CStr(paragraphRange.End) & vbLf & _
        "caret=" & CStr(Selection.Start) & vbLf
    Exit Sub

SetupFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "stage=setup" & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX physical Return regression setup", _
        regressionErrorDescription
End Sub

Public Sub VisualTeX_InspectWordImagePhysicalReturnRegression()
    Const formulaId As String = _
        "29292929-2929-4929-8929-292929292929"

    Dim testDocument As Document
    Dim formulaRange As Range
    Dim formulaParagraph As Range
    Dim helperParagraph As Range
    Dim typingParagraph As Range
    Dim sequenceField As Field
    Dim visibleNumberField As Field
    Dim numberRange As Range
    Dim resultPath As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo InspectFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-image-physical-return-regression-result.txt"
    If Documents.Count <> 1 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The physical Return regression document is not the only open document."
    End If
    Set testDocument = ActiveDocument
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    If formulaRange Is Nothing Or formulaRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The physical Return regression lost its numbered image formula."
    End If
    Set formulaParagraph = formulaRange.Paragraphs(1).Range.Duplicate
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, formulaId)
    Set visibleNumberField = VTImageEquationReferenceField( _
        formulaRange, formulaId)
    If Not visibleNumberField Is Nothing Then
        Set numberRange = VTImageEquationNumberRange( _
            formulaRange, visibleNumberField)
    End If
    If sequenceField Is Nothing Or visibleNumberField Is Nothing Or _
       numberRange Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The physical Return regression lost its external SEQ or visible REF."
    End If
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    Set typingParagraph = Selection.Range.Paragraphs(1).Range.Duplicate

    If helperParagraph.Start <> formulaParagraph.End Or _
       typingParagraph.Start <> helperParagraph.End Or _
       typingParagraph.Information(wdWithInTable) Or _
       typingParagraph.Fields.Count <> 0 Or _
       typingParagraph.InlineShapes.Count <> 0 Or _
       typingParagraph.OMaths.Count <> 0 Or _
       VTCustomTabStopCount(typingParagraph) <> 0 Or _
       typingParagraph.Font.Position <> 0 Or _
       VTWordRangeHasMeaningfulText(typingParagraph) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Physical Return did not end as formula -> helper -> ordinary typing line" & _
            " [formula=" & CStr(formulaParagraph.Start) & "-" & _
                CStr(formulaParagraph.End) & _
            "; helper=" & CStr(helperParagraph.Start) & "-" & _
                CStr(helperParagraph.End) & _
            "; typing=" & CStr(typingParagraph.Start) & "-" & _
                CStr(typingParagraph.End) & "]."
    End If

    VTWriteTextAtomic _
        resultPath, _
        "PASS" & vbLf & _
        "structure=<TAB><image><TAB>(REF)<CR>+SEQ-helper+typing-line" & vbLf & _
        "formula=" & CStr(formulaParagraph.Start) & "-" & _
            CStr(formulaParagraph.End) & vbLf & _
        "helper=" & CStr(helperParagraph.Start) & "-" & _
            CStr(helperParagraph.End) & vbLf & _
        "typing=" & CStr(typingParagraph.Start) & "-" & _
            CStr(typingParagraph.End) & vbLf & _
        "selection=" & CStr(Selection.Start) & vbLf
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Exit Sub

InspectFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "stage=inspect-after-physical-return" & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX physical Return regression inspect", _
        regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordRealImagePaintedCenterRegression()
    Dim testDocument As Document
    Dim formulaIds As Variant
    Dim formulaWidths As Variant
    Dim formulaHeights As Variant
    Dim inkCenterRatios As Variant
    Dim expectedPositions As Variant
    Dim formulaId As String
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim insertionRange As Range
    Dim visibleNumberField As Field
    Dim numberRange As Range
    Dim encodedMetadata As String
    Dim resultPath As String
    Dim report As String
    Dim centerDeltaPoints As Double
    Dim positionMagnitude As Long
    Dim actualPosition As Long
    Dim itemIndex As Long
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-real-image-painted-center-regression-result.txt"
    formulaIds = Array( _
        "b28e7d9a-3a62-4e62-8463-e5308ea20e43", _
        "a6efab1e-830c-4888-8c54-34e0a5f06d4d", _
        "36b9dfd1-7de0-46a7-ac7f-142048da394d")
    formulaWidths = Array( _
        64.2129921259843, 135.343858267717, 236.453464566929)
    formulaHeights = Array( _
        12.9820472440945, 33.4684251968504, 33.6983464566929)
    inkCenterRatios = Array(0.5, 0.49122807, 0.47826087)
    expectedPositions = Array(-4, -15, -15)

    If Documents.Count > 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The real-image painted-center regression requires Word to have no open documents."
    End If
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.Activate
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Styles(wdStyleNormal).Font.Size = 10!
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_MODE_VARIABLE, _
        VT_WORD_NUMBERING_MODE_SEQUENCE
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_SEPARATOR_VARIABLE, "."
    encodedMetadata = VT_METADATA_PREFIX & "e30"
    report = "PASS" & vbLf & _
        "alignment=cached-PNG-ink-center+Cambria-Math-glyph-center" & vbLf & _
        "numberGlyphCenterAboveBaselineRatio=" & _
            CStr(VT_WORD_EQUATION_NUMBER_INK_CENTER_ABOVE_BASELINE_RATIO) & vbLf

    For itemIndex = LBound(formulaIds) To UBound(formulaIds)
        formulaId = CStr(formulaIds(itemIndex))
        regressionStage = "insert-real-image-" & CStr(itemIndex + 1)
        Set insertionRange = VTAppendRegressionParagraph(testDocument)
        Set formulaShape = testDocument.InlineShapes.AddPicture( _
            FileName:=VTImageCachePngPath(formulaId), _
            LinkToFile:=False, SaveWithDocument:=True, _
            Range:=insertionRange)
        formulaShape.LockAspectRatio = msoFalse
        formulaShape.Width = CSng(formulaWidths(itemIndex))
        formulaShape.Height = CSng(formulaHeights(itemIndex))
        formulaShape.LockAspectRatio = msoTrue
        formulaShape.AlternativeText = encodedMetadata
        formulaShape.Title = VTFormulaReference(formulaId, "block", True)
        VTSetWordLatexPayload testDocument, formulaId, "eF8y"
        VTSetWordMetadataPayload testDocument, formulaId, encodedMetadata
        VTSetWordFormulaFormat testDocument, formulaId, "block", True
        VTSetWordImageInkCenterYRatio _
            testDocument, formulaId, CDbl(inkCenterRatios(itemIndex))
        Set formulaRange = VTInsertEquationNumber( _
            formulaShape, formulaId, "real image painted-center fixture")

        regressionStage = "measure-real-image-" & CStr(itemIndex + 1)
        Set formulaRange = VTNumberedFormulaRangeForId( _
            testDocument, formulaId)
        Set formulaShape = formulaRange.InlineShapes(1)
        Set visibleNumberField = VTImageEquationReferenceField( _
            formulaRange, formulaId)
        If Not visibleNumberField Is Nothing Then
            Set numberRange = VTImageEquationNumberRange( _
                formulaRange, visibleNumberField)
        Else
            Set numberRange = Nothing
        End If
        If visibleNumberField Is Nothing Or numberRange Is Nothing Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The real-image fixture lost its visible Equation REF."
        End If
        testDocument.Repaginate
        ActiveWindow.ScrollIntoView formulaRange, True
        DoEvents
        positionMagnitude = VTApplyNumberedImageFormulaVerticalAlignment( _
            formulaShape, numberRange, _
            VTVisibleEquationNumberFontSize(testDocument), _
            "The real-image painted-center correction")
        testDocument.Repaginate
        DoEvents
        actualPosition = formulaShape.Range.Font.Position
        If actualPosition <> CLng(expectedPositions(itemIndex)) Or _
           positionMagnitude <> Abs(CLng(expectedPositions(itemIndex))) Or _
           numberRange.Font.Position <> 0 Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The real-image fixture did not apply the screen-verified picture-only Position" & _
                " [index=" & CStr(itemIndex + 1) & _
                "; expected=" & CStr(expectedPositions(itemIndex)) & _
                "; actual=" & CStr(actualPosition) & _
                "; number=" & CStr(numberRange.Font.Position) & "]."
        End If
        If Not VTTryCalculateCurrentNumberedImagePaintedCenterDeltaPoints( _
           formulaShape, numberRange, centerDeltaPoints) Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "Word did not expose real-image ink-center geometry."
        End If
        If Abs(centerDeltaPoints) > 0.5001# Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The real-image ink centers exceed Position quantisation" & _
                " [index=" & CStr(itemIndex + 1) & _
                "; deltaPoints=" & CStr(centerDeltaPoints) & _
                "; position=" & CStr(formulaShape.Range.Font.Position) & "]."
        End If
        report = report & _
            "formula" & CStr(itemIndex + 1) & _
            ".size=" & CStr(formulaShape.Width) & "x" & _
                CStr(formulaShape.Height) & vbLf & _
            "formula" & CStr(itemIndex + 1) & _
            ".inkCenter=" & CStr(inkCenterRatios(itemIndex)) & vbLf & _
            "formula" & CStr(itemIndex + 1) & _
            ".position=" & CStr(actualPosition) & vbLf & _
            "formula" & CStr(itemIndex + 1) & _
            ".expectedPosition=" & CStr(expectedPositions(itemIndex)) & vbLf & _
            "formula" & CStr(itemIndex + 1) & _
            ".positionMagnitude=" & CStr(positionMagnitude) & vbLf & _
            "formula" & CStr(itemIndex + 1) & _
            ".deltaPoints=" & CStr(centerDeltaPoints) & vbLf
    Next itemIndex

    regressionStage = "write-real-image-result"
    VTWriteTextAtomic resultPath, report
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    ' Preserve the failed document and its exact Word state for inspection.
    ' The test driver decides whether it is safe to close or restart Word only
    ' after recording dialogs, processes, documents, windows and result files.
    If Not testDocument Is Nothing Then testDocument.Activate
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX real-image painted-center regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordImagePositionGetPointProbe()
    Const formulaId As String = _
        "36b9dfd1-7de0-46a7-ac7f-142048da394d"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim insertionRange As Range
    Dim probeRange As Range
    Dim positions As Variant
    Dim itemIndex As Long
    Dim positionValue As Long
    Dim leftPixels As Long
    Dim topPixels As Long
    Dim widthPixels As Long
    Dim heightPixels As Long
    Dim verticalInfo As Variant
    Dim resultPath As String
    Dim report As String
    Dim probeErrorNumber As Long
    Dim probeErrorDescription As String

    On Error GoTo ProbeFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-image-position-getpoint-probe-result.txt"
    If Documents.Count > 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The image Position/GetPoint probe requires Word to have no open documents."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.Activate
    testDocument.ActiveWindow.View.Type = wdPrintView
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    Set formulaShape = testDocument.InlineShapes.AddPicture( _
        FileName:=VTImageCachePngPath(formulaId), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    formulaShape.LockAspectRatio = msoFalse
    formulaShape.Width = 236.45346!
    formulaShape.Height = 33.69835!
    formulaShape.LockAspectRatio = msoTrue
    Set probeRange = formulaShape.Range.Duplicate
    positions = Array(-2, -1, 0, 1, 2)
    report = "DIAGNOSTIC_ONLY" & vbLf & _
        "conclusion=GetPoint-is-a-row-box-and-not-a-painted-bbox" & vbLf

    For itemIndex = LBound(positions) To UBound(positions)
        positionValue = CLng(positions(itemIndex))
        formulaShape.Range.Font.Position = positionValue
        DoEvents
        testDocument.Repaginate
        ActiveWindow.GetPoint _
            leftPixels, topPixels, widthPixels, heightPixels, probeRange
        verticalInfo = probeRange.Information( _
            wdVerticalPositionRelativeToPage)
        report = report & _
            "position=" & CStr(positionValue) & _
            "; top=" & CStr(topPixels) & _
            "; height=" & CStr(heightPixels) & _
            "; infoY=" & CStr(verticalInfo) & vbLf
    Next itemIndex

    VTWriteTextAtomic resultPath, report
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Exit Sub

ProbeFailed:
    probeErrorNumber = Err.Number
    probeErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "errorNumber=" & CStr(probeErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(probeErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    If Not testDocument Is Nothing Then
        testDocument.Saved = True
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise probeErrorNumber, _
        "VisualTeX image Position/GetPoint probe", _
        probeErrorDescription
End Sub

Public Sub VisualTeX_SetupWordRealImageScreenCenterProbe()
    Dim testDocument As Document
    Dim formulaIds As Variant
    Dim formulaWidths As Variant
    Dim formulaHeights As Variant
    Dim inkCenterRatios As Variant
    Dim formulaId As String
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim insertionRange As Range
    Dim visibleNumberField As Field
    Dim numberRange As Range
    Dim encodedMetadata As String
    Dim resultPath As String
    Dim report As String
    Dim itemIndex As Long
    Dim formulaLeft As Long
    Dim formulaTop As Long
    Dim formulaWidthPixels As Long
    Dim formulaHeightPixels As Long
    Dim numberLeft As Long
    Dim numberTop As Long
    Dim numberWidthPixels As Long
    Dim numberHeightPixels As Long
    Dim probeErrorNumber As Long
    Dim probeErrorDescription As String

    On Error GoTo ProbeFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-real-image-screen-center-layout.txt"
    If Documents.Count > 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The real-image screen-center probe requires Word to have no open documents."
    End If

    formulaIds = Array( _
        "b28e7d9a-3a62-4e62-8463-e5308ea20e43", _
        "a6efab1e-830c-4888-8c54-34e0a5f06d4d", _
        "36b9dfd1-7de0-46a7-ac7f-142048da394d")
    formulaWidths = Array( _
        64.2129921259843, 135.343858267717, 236.453464566929)
    formulaHeights = Array( _
        12.9820472440945, 33.4684251968504, 33.6983464566929)
    inkCenterRatios = Array(0.5, 0.49122807, 0.47826087)
    encodedMetadata = VT_METADATA_PREFIX & "e30"

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.Activate
    testDocument.ActiveWindow.View.Type = wdPrintView
    On Error Resume Next
    testDocument.ActiveWindow.View.Zoom.Percentage = 100
    On Error GoTo ProbeFailed
    testDocument.Styles(wdStyleNormal).Font.Size = 10!
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_MODE_VARIABLE, _
        VT_WORD_NUMBERING_MODE_SEQUENCE
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_SEPARATOR_VARIABLE, "."
    report = "READY" & vbLf & _
        "getPointMeaning=row-box-crop-hint-only" & vbLf & _
        "acceptance=external-screen-pixel-painted-bbox" & vbLf

    For itemIndex = LBound(formulaIds) To UBound(formulaIds)
        formulaId = CStr(formulaIds(itemIndex))
        Set insertionRange = VTAppendRegressionParagraph(testDocument)
        Set formulaShape = testDocument.InlineShapes.AddPicture( _
            FileName:=VTImageCachePngPath(formulaId), _
            LinkToFile:=False, SaveWithDocument:=True, _
            Range:=insertionRange)
        formulaShape.LockAspectRatio = msoFalse
        formulaShape.Width = CSng(formulaWidths(itemIndex))
        formulaShape.Height = CSng(formulaHeights(itemIndex))
        formulaShape.LockAspectRatio = msoTrue
        formulaShape.AlternativeText = encodedMetadata
        formulaShape.Title = VTFormulaReference(formulaId, "block", True)
        VTSetWordLatexPayload testDocument, formulaId, "eF8y"
        VTSetWordMetadataPayload testDocument, formulaId, encodedMetadata
        VTSetWordFormulaFormat testDocument, formulaId, "block", True
        VTSetWordImageInkCenterYRatio _
            testDocument, formulaId, CDbl(inkCenterRatios(itemIndex))
        Set formulaRange = VTInsertEquationNumber( _
            formulaShape, formulaId, "real image screen-center fixture")
        Set formulaRange = VTNumberedFormulaRangeForId( _
            testDocument, formulaId)
        Set formulaShape = formulaRange.InlineShapes(1)
        Set visibleNumberField = VTImageEquationReferenceField( _
            formulaRange, formulaId)
        If visibleNumberField Is Nothing Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The screen-center fixture lost its visible REF."
        End If
        Set numberRange = VTImageEquationNumberRange( _
            formulaRange, visibleNumberField)
        If numberRange Is Nothing Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The screen-center fixture lost its visible number range."
        End If

        testDocument.Repaginate
        ActiveWindow.ScrollIntoView formulaRange, True
        DoEvents
        ActiveWindow.GetPoint _
            formulaLeft, formulaTop, formulaWidthPixels, formulaHeightPixels, _
            formulaShape.Range
        ActiveWindow.GetPoint _
            numberLeft, numberTop, numberWidthPixels, numberHeightPixels, _
            numberRange
        report = report & _
            "formula" & CStr(itemIndex + 1) & _
            ".box=" & CStr(formulaLeft) & "," & CStr(formulaTop) & "," & _
                CStr(formulaWidthPixels) & "," & CStr(formulaHeightPixels) & vbLf & _
            "formula" & CStr(itemIndex + 1) & _
            ".numberBox=" & CStr(numberLeft) & "," & CStr(numberTop) & "," & _
                CStr(numberWidthPixels) & "," & CStr(numberHeightPixels) & vbLf & _
            "formula" & CStr(itemIndex + 1) & _
            ".position=" & CStr(formulaShape.Range.Font.Position) & vbLf
    Next itemIndex

    testDocument.Range( _
        Start:=testDocument.Content.End - 1, _
        End:=testDocument.Content.End - 1).Select
    testDocument.ActiveWindow.ScrollIntoView _
        testDocument.Range(Start:=0, End:=1), True
    DoEvents

    ' GetPoint is viewport-relative and reports Word's row box, not the image or
    ' number painted bbox; it is only a crop hint for an external screen-pixel ink
    ' scan. Re-read it after the final view is fixed so that hint is coherent.
    report = "READY" & vbLf & _
        "getPointMeaning=row-box-crop-hint-only" & vbLf & _
        "acceptance=external-screen-pixel-painted-bbox" & vbLf
    For itemIndex = LBound(formulaIds) To UBound(formulaIds)
        formulaId = CStr(formulaIds(itemIndex))
        Set formulaRange = VTNumberedFormulaRangeForId( _
            testDocument, formulaId)
        Set formulaShape = formulaRange.InlineShapes(1)
        Set visibleNumberField = VTImageEquationReferenceField( _
            formulaRange, formulaId)
        Set numberRange = VTImageEquationNumberRange( _
            formulaRange, visibleNumberField)
        ActiveWindow.GetPoint _
            formulaLeft, formulaTop, formulaWidthPixels, formulaHeightPixels, _
            formulaShape.Range
        ActiveWindow.GetPoint _
            numberLeft, numberTop, numberWidthPixels, numberHeightPixels, _
            numberRange
        report = report & _
            "formula" & CStr(itemIndex + 1) & _
            ".box=" & CStr(formulaLeft) & "," & CStr(formulaTop) & "," & _
                CStr(formulaWidthPixels) & "," & CStr(formulaHeightPixels) & vbLf & _
            "formula" & CStr(itemIndex + 1) & _
            ".numberBox=" & CStr(numberLeft) & "," & CStr(numberTop) & "," & _
                CStr(numberWidthPixels) & "," & CStr(numberHeightPixels) & vbLf & _
            "formula" & CStr(itemIndex + 1) & _
            ".position=" & CStr(formulaShape.Range.Font.Position) & vbLf
    Next itemIndex
    VTWriteTextAtomic resultPath, report
    Exit Sub

ProbeFailed:
    probeErrorNumber = Err.Number
    probeErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        resultPath, _
        "FAIL" & vbLf & _
        "errorNumber=" & CStr(probeErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(probeErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    On Error GoTo 0
    Err.Raise probeErrorNumber, _
        "VisualTeX real-image screen-center probe", _
        probeErrorDescription
End Sub

Private Sub VTFocusWordRealImageScreenCenterProbe( _
    ByVal formulaId As String, _
    ByVal itemIndex As Long)

    Dim testDocument As Document
    Dim formulaRange As Range
    Dim formulaShape As InlineShape
    Dim visibleNumberField As Field
    Dim numberRange As Range
    Dim caretRange As Range
    Dim formulaLeft As Long
    Dim formulaTop As Long
    Dim formulaWidthPixels As Long
    Dim formulaHeightPixels As Long
    Dim numberLeft As Long
    Dim numberTop As Long
    Dim numberWidthPixels As Long
    Dim numberHeightPixels As Long
    Dim resultPath As String

    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-real-image-screen-focus.txt"
    If Documents.Count <> 1 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The real-image screen focus probe requires exactly one open fixture document."
    End If
    Set testDocument = ActiveDocument
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    If formulaRange Is Nothing Or formulaRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The requested real-image screen fixture could not be resolved."
    End If
    Set formulaShape = formulaRange.InlineShapes(1)
    Set visibleNumberField = VTImageEquationReferenceField( _
        formulaRange, formulaId)
    If visibleNumberField Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The requested real-image screen fixture lost its visible REF."
    End If
    Set numberRange = VTImageEquationNumberRange( _
        formulaRange, visibleNumberField)
    If numberRange Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The requested real-image screen fixture lost its number range."
    End If

    ActiveWindow.ScrollIntoView formulaRange, True
    DoEvents
    Set caretRange = testDocument.Range( _
        Start:=formulaRange.Paragraphs(1).Range.End - 1, _
        End:=formulaRange.Paragraphs(1).Range.End - 1)
    caretRange.Select
    DoEvents
    ActiveWindow.GetPoint _
        formulaLeft, formulaTop, formulaWidthPixels, formulaHeightPixels, _
        formulaShape.Range
    ActiveWindow.GetPoint _
        numberLeft, numberTop, numberWidthPixels, numberHeightPixels, _
        numberRange
    VTWriteTextAtomic _
        resultPath, _
        "READY" & vbLf & _
        "index=" & CStr(itemIndex) & vbLf & _
        "box=" & CStr(formulaLeft) & "," & CStr(formulaTop) & "," & _
            CStr(formulaWidthPixels) & "," & CStr(formulaHeightPixels) & vbLf & _
        "numberBox=" & CStr(numberLeft) & "," & CStr(numberTop) & "," & _
            CStr(numberWidthPixels) & "," & CStr(numberHeightPixels) & vbLf & _
        "position=" & CStr(formulaShape.Range.Font.Position) & vbLf
End Sub

Public Sub VisualTeX_FocusWordRealImageScreenCenterProbe1()
    VTFocusWordRealImageScreenCenterProbe _
        "b28e7d9a-3a62-4e62-8463-e5308ea20e43", 1
End Sub

Public Sub VisualTeX_FocusWordRealImageScreenCenterProbe2()
    VTFocusWordRealImageScreenCenterProbe _
        "a6efab1e-830c-4888-8c54-34e0a5f06d4d", 2
End Sub

Public Sub VisualTeX_FocusWordRealImageScreenCenterProbe3()
    VTFocusWordRealImageScreenCenterProbe _
        "36b9dfd1-7de0-46a7-ac7f-142048da394d", 3
End Sub

Private Function VTRegressionCreateNumberedImageAtCaret( _
    ByVal documentObject As Document, _
    ByVal requestedRange As Range, _
    ByVal formulaId As String, _
    ByVal latexBase64 As String, _
    ByVal widthPoints As Single, _
    ByVal heightPoints As Single) As InlineShape

    Dim insertionRange As Range
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim encodedMetadata As String

    If documentObject Is Nothing Or requestedRange Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The numbering performance insertion target is invalid."
    End If
    encodedMetadata = VT_METADATA_PREFIX & "e30"
    Set insertionRange = VTPrepareWordCreateInsertionRange( _
        requestedRange.Duplicate, "block")
    Set formulaShape = documentObject.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    formulaShape.Width = widthPoints
    formulaShape.Height = heightPoints
    formulaShape.AlternativeText = encodedMetadata
    formulaShape.Title = VTFormulaReference(formulaId, "block", True)
    VTSetWordLatexPayload documentObject, formulaId, latexBase64
    VTSetWordMetadataPayload documentObject, formulaId, encodedMetadata
    VTSetWordFormulaFormat documentObject, formulaId, "block", True
    Set formulaRange = VTInsertEquationNumber( _
        formulaShape, formulaId, _
        VTEquationCrossReferenceText(latexBase64))
    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If formulaRange Is Nothing Or _
       formulaRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The numbering performance insertion lost its image formula."
    End If
    Set VTRegressionCreateNumberedImageAtCaret = _
        formulaRange.InlineShapes(1)
End Function

Private Function VTRegressionCreateDirectSeqImageAtRange( _
    ByVal documentObject As Document, _
    ByVal requestedRange As Range, _
    ByVal formulaId As String) As InlineShape

    Dim insertionRange As Range
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim formulaContainerRange As Range
    Dim paragraphRange As Range
    Dim fieldRange As Range
    Dim sequenceField As Field
    Dim numberRange As Range
    Dim captionRange As Range
    Dim sequenceBookmarkName As String
    Dim captionBookmarkName As String
    Dim fieldCode As String
    Dim paragraphStart As Long

    If documentObject Is Nothing Or requestedRange Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The direct-SEQ image prototype target is invalid."
    End If

    Set insertionRange = VTPrepareWordCreateInsertionRange( _
        requestedRange.Duplicate, "block")
    Set formulaShape = documentObject.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    formulaShape.Width = 90!
    formulaShape.Height = 24!
    formulaShape.AlternativeText = VT_METADATA_PREFIX & "e30"
    formulaShape.Title = VTFormulaReference(formulaId, "block", True)
    paragraphStart = formulaShape.Range.Paragraphs(1).Range.Start

    Set formulaRange = _
        VTPrependCenterTabPreservingImage(formulaShape.Range.Duplicate)
    Set formulaShape = formulaRange.InlineShapes(1)
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    VTConfigureNumberedEquationParagraph paragraphRange

    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, paragraphStart)
    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set insertionRange = documentObject.Range( _
        Start:=formulaContainerRange.End, End:=formulaContainerRange.End)
    insertionRange.Text = vbTab & "()"

    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, paragraphStart)
    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set fieldRange = documentObject.Range( _
        Start:=formulaContainerRange.End + 2, _
        End:=formulaContainerRange.End + 2)
    fieldCode = "SEQ " & _
        VTEquationSequenceFieldText(VTNativeEquationLabelName())
    Set sequenceField = fieldRange.Fields.Add( _
        Range:=fieldRange, Type:=wdFieldEmpty, _
        Text:=fieldCode, PreserveFormatting:=True)
    sequenceField.Update

    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, paragraphStart)
    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    Set numberRange = documentObject.Range( _
        Start:=formulaContainerRange.End + 1, _
        End:=paragraphRange.End - 1)
    If numberRange.Fields.Count <> 1 Or _
       Left$(numberRange.Text, 1) <> "(" Or _
       Right$(numberRange.Text, 1) <> ")" Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not keep the direct SEQ inside the visible number label."
    End If
    VTApplyStaticImageEquationNumberFormatting _
        numberRange, VTVisibleEquationNumberFontSize(documentObject)

    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then _
        documentObject.Bookmarks(sequenceBookmarkName).Delete
    documentObject.Bookmarks.Add _
        Name:=sequenceBookmarkName, Range:=sequenceField.Result.Duplicate
    VTSetEquationNumberBookmarkExact _
        documentObject, formulaId, numberRange
    Set captionRange = numberRange.Duplicate
    captionRange.Collapse wdCollapseEnd
    If documentObject.Bookmarks.Exists(captionBookmarkName) Then _
        documentObject.Bookmarks(captionBookmarkName).Delete
    documentObject.Bookmarks.Add _
        Name:=captionBookmarkName, Range:=captionRange

    Set VTRegressionCreateDirectSeqImageAtRange = formulaShape
End Function

Public Sub VisualTeX_RunWordDirectSeqImagePrototype()
    Const formulaCount As Long = 3

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim formulaIds(1 To formulaCount) As String
    Dim insertionRange As Range
    Dim formulaRange As Range
    Dim sequenceField As Field
    Dim candidateField As Field
    Dim referenceField As Field
    Dim referenceRange As Range
    Dim blankParagraph As Range
    Dim items As Variant
    Dim builtInField As Field
    Dim builtInTargetName As String
    Dim builtInResult As String
    Dim bodyFieldCountBefore As Long
    Dim bodyFieldCountAfter As Long
    Dim resultPath As String
    Dim resultText As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim itemIndex As Long
    Dim sequenceCount As Long
    Dim fieldStart As Long

    On Error GoTo PrototypeFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-direct-seq-image-prototype-result.txt"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate

    regressionStage = "create-three-direct-seq-images"
    For itemIndex = 1 To formulaCount
        formulaIds(itemIndex) = _
            VTRegressionPerformanceFormulaId(930000000 + itemIndex)
        Set insertionRange = VTAppendRegressionParagraph(testDocument)
        Set formulaShape = VTRegressionCreateDirectSeqImageAtRange( _
            testDocument, insertionRange, formulaIds(itemIndex))
    Next itemIndex
    testDocument.Fields.Update

    regressionStage = "verify-direct-seq-results"
    For itemIndex = 1 To formulaCount
        Set sequenceField = VTEquationSequenceFieldForBookmark( _
            testDocument, _
            VTEquationSequenceNumberBookmarkName(formulaIds(itemIndex)))
        If sequenceField Is Nothing Or _
           Trim$(sequenceField.Result.Text) <> CStr(itemIndex) Or _
           sequenceField.Result.Paragraphs(1).Range.InlineShapes.Count <> 1 Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "A direct image SEQ did not remain inside its visible formula row."
        End If
    Next itemIndex
    For Each candidateField In testDocument.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            sequenceCount = sequenceCount + 1
        End If
    Next candidateField
    If sequenceCount <> formulaCount Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The direct image prototype created an unexpected SEQ inventory."
    End If

    regressionStage = "body-ref-to-direct-seq"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set referenceRange = insertionRange.Duplicate
    Set referenceField = referenceRange.Fields.Add( _
        Range:=referenceRange, Type:=wdFieldEmpty, _
        Text:="REF " & _
            VTEquationSequenceNumberBookmarkName(formulaIds(2)) & " \\h", _
        PreserveFormatting:=True)
    referenceField.Update
    If Trim$(referenceField.Result.Text) <> "2" Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "A body REF could not target the direct image SEQ Bookmark."
    End If

    regressionStage = "word-built-in-cross-reference-list"
    items = testDocument.GetCrossReferenceItems(wdCaptionEquation)
    If Not IsArray(items) Or UBound(items) < formulaCount Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not expose the direct image SEQs in its Equation list."
    End If

    regressionStage = "word-built-in-cross-reference-insert"
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            bodyFieldCountBefore = bodyFieldCountBefore + 1
        End If
    Next candidateField
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    insertionRange.Select
    Selection.InsertCrossReference _
        ReferenceType:=wdCaptionEquation, _
        ReferenceKind:=wdOnlyLabelAndNumber, _
        ReferenceItem:=2, _
        InsertAsHyperlink:=True, _
        IncludePosition:=False
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            bodyFieldCountAfter = bodyFieldCountAfter + 1
            If bodyFieldCountAfter > bodyFieldCountBefore Then
                Set builtInField = candidateField
            End If
        End If
    Next candidateField
    If bodyFieldCountAfter <> bodyFieldCountBefore + 1 Or _
       builtInField Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not insert exactly one built-in Equation cross-reference."
    End If
    builtInTargetName = VTReferenceTargetBookmarkName( _
        builtInField.Code.Text)
    builtInResult = VTComparableEquationNumberText( _
        builtInField.Result.Text)
    If builtInResult <> "2" Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word built-in cross-reference to direct SEQ is not a pure number" & _
            " [target=" & builtInTargetName & _
            "; result=" & builtInField.Result.Text & "]."
    End If

    regressionStage = "return-after-direct-seq-row"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaIds(2))
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The second direct image formula could not be resolved."
    End If
    Set insertionRange = testDocument.Range( _
        Start:=formulaRange.Paragraphs(1).Range.End - 1, _
        End:=formulaRange.Paragraphs(1).Range.End - 1)
    insertionRange.Select
    Selection.TypeParagraph
    DoEvents
    Set blankParagraph = Selection.Range.Paragraphs(1).Range.Duplicate
    If blankParagraph.Fields.Count <> 0 Or _
       blankParagraph.InlineShapes.Count <> 0 Or _
       blankParagraph.OMaths.Count <> 0 Or _
       VTCustomTabStopCount(blankParagraph) <> 0 Or _
       VTWordRangeHasMeaningfulText(blankParagraph) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Return after a direct image SEQ did not create a plain paragraph."
    End If
    testDocument.Fields.Update
    For itemIndex = 1 To formulaCount
        Set sequenceField = VTEquationSequenceFieldForBookmark( _
            testDocument, _
            VTEquationSequenceNumberBookmarkName(formulaIds(itemIndex)))
        If sequenceField Is Nothing Or _
           Trim$(sequenceField.Result.Text) <> CStr(itemIndex) Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "Return changed direct image SEQ ordering."
        End If
    Next itemIndex

    regressionStage = "insert-direct-seq-in-new-line"
    formulaIds(3) = VTRegressionPerformanceFormulaId(930000004)
    Set formulaShape = VTRegressionCreateDirectSeqImageAtRange( _
        testDocument, blankParagraph, formulaIds(3))
    testDocument.Fields.Update
    sequenceCount = 0
    For Each candidateField In testDocument.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            sequenceCount = sequenceCount + 1
        End If
    Next candidateField
    If sequenceCount <> 4 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Middle direct image insertion did not create exactly one new SEQ."
    End If

    resultText = "PASS" & vbLf & _
        "structure=<TAB><image><TAB>(SEQ)<CR>" & vbLf & _
        "externalHelperParagraphs=0" & vbLf & _
        "fieldUpdate=PASS" & vbLf & _
        "bodyRef=PASS" & vbLf & _
        "wordEquationList=PASS" & vbLf & _
        "wordBuiltInCrossReference=PASS" & vbLf & _
        "wordBuiltInTarget=" & builtInTargetName & vbLf & _
        "return=PASS" & vbLf & _
        "middleInsert=PASS" & vbLf
    VTWriteTextAtomic resultPath, resultText
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

PrototypeFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    resultText = "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    VTWriteTextAtomic resultPath, resultText
    If Not testDocument Is Nothing Then
        testDocument.Saved = True
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word direct image SEQ prototype", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Private Function VTRegressionCreateEmbeddedSeqImageAtRange( _
    ByVal documentObject As Document, _
    ByVal requestedRange As Range, _
    ByVal formulaId As String) As InlineShape

    Dim insertionRange As Range
    Dim formulaShape As InlineShape
    Dim paragraphRange As Range
    Dim formulaContainerRange As Range
    Dim fieldRange As Range
    Dim sequenceField As Field
    Dim visibleRef As Field
    Dim numberRange As Range
    Dim captionRange As Range
    Dim sequenceBookmarkName As String
    Dim captionBookmarkName As String
    Dim paragraphStart As Long

    If documentObject Is Nothing Or requestedRange Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The embedded-SEQ image prototype target is invalid."
    End If

    Set insertionRange = VTPrepareWordCreateInsertionRange( _
        requestedRange.Duplicate, "block")
    Set formulaShape = documentObject.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    formulaShape.Width = 90!
    formulaShape.Height = 24!
    formulaShape.AlternativeText = VT_METADATA_PREFIX & "e30"
    formulaShape.Title = VTFormulaReference(formulaId, "block", True)
    paragraphStart = formulaShape.Range.Paragraphs(1).Range.Start
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    VTConfigureNumberedEquationParagraph paragraphRange

    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)

    ' Keep the true Equation SEQ in the same visible formula paragraph, but at
    ' its very beginning and hidden from layout. Word's caption machinery then
    ' sees the SEQ before any tab/image content while Return can never displace
    ' the sequence source into another paragraph.
    Set fieldRange = documentObject.Range( _
        Start:=paragraphStart, End:=paragraphStart)
    Set sequenceField = fieldRange.Fields.Add( _
        Range:=fieldRange, Type:=wdFieldEmpty, _
        Text:="SEQ " & VTEquationSequenceFieldText( _
            VTNativeEquationLabelName()), _
        PreserveFormatting:=False)
    sequenceField.Update
    With sequenceField.Result.Font
        .Hidden = True
        .Color = wdColorWhite
        .Size = 1!
        .Position = 0
    End With
    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then _
        documentObject.Bookmarks(sequenceBookmarkName).Delete
    documentObject.Bookmarks.Add _
        Name:=sequenceBookmarkName, Range:=sequenceField.Result.Duplicate

    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, paragraphStart)
    Set insertionRange = documentObject.Range( _
        Start:=formulaShape.Range.Start, End:=formulaShape.Range.Start)
    insertionRange.Text = vbTab
    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, paragraphStart)
    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set insertionRange = documentObject.Range( _
        Start:=formulaContainerRange.End, End:=formulaContainerRange.End)
    insertionRange.Text = vbTab & "()"

    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, paragraphStart)
    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set fieldRange = documentObject.Range( _
        Start:=formulaContainerRange.End + 2, _
        End:=formulaContainerRange.End + 2)
    Set visibleRef = fieldRange.Fields.Add( _
        Range:=fieldRange, Type:=wdFieldRef, _
        Text:=sequenceBookmarkName, PreserveFormatting:=False)
    visibleRef.Update
    With visibleRef.Result.Font
        .Hidden = False
        .Color = wdColorAutomatic
        .Name = VT_WORD_EQUATION_NUMBER_FONT_NAME
        .NameAscii = VT_WORD_EQUATION_NUMBER_FONT_NAME
        .NameOther = VT_WORD_EQUATION_NUMBER_FONT_NAME
        .Position = 0
        .Size = VTVisibleEquationNumberFontSize(documentObject)
    End With

    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, paragraphStart)
    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    Set numberRange = documentObject.Range( _
        Start:=formulaContainerRange.End + 1, _
        End:=paragraphRange.End - 1)
    If numberRange.Fields.Count <> 1 Or _
       Left$(numberRange.Text, 1) <> "(" Or _
       Right$(numberRange.Text, 1) <> ")" Or _
       VTReferenceTargetBookmarkName( _
           numberRange.Fields(1).Code.Text) <> sequenceBookmarkName Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not keep the visible REF beside the embedded SEQ image."
    End If
    VTSetEquationNumberBookmarkExact _
        documentObject, formulaId, numberRange
    Set captionRange = sequenceField.Result.Duplicate
    If documentObject.Bookmarks.Exists(captionBookmarkName) Then _
        documentObject.Bookmarks(captionBookmarkName).Delete
    documentObject.Bookmarks.Add _
        Name:=captionBookmarkName, Range:=captionRange

    Set VTRegressionCreateEmbeddedSeqImageAtRange = formulaShape
End Function

Public Sub VisualTeX_RunWordEmbeddedSeqImagePrototype()
    Const formulaCount As Long = 3

    Dim testDocument As Document
    Dim formulaIds(1 To formulaCount) As String
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim insertionRange As Range
    Dim blankParagraph As Range
    Dim sequenceField As Field
    Dim visibleRef As Field
    Dim candidateField As Field
    Dim builtInField As Field
    Dim items As Variant
    Dim resultPath As String
    Dim resultText As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim builtInTargetName As String
    Dim builtInResult As String
    Dim itemIndex As Long
    Dim sequenceCount As Long
    Dim bodyFieldCountBefore As Long
    Dim bodyFieldCountAfter As Long

    On Error GoTo PrototypeFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-embedded-seq-image-prototype-result.txt"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate

    regressionStage = "create-three-embedded-seq-images"
    For itemIndex = 1 To formulaCount
        formulaIds(itemIndex) = _
            VTRegressionPerformanceFormulaId(940000000 + itemIndex)
        Set insertionRange = VTAppendRegressionParagraph(testDocument)
        Set formulaShape = VTRegressionCreateEmbeddedSeqImageAtRange( _
            testDocument, insertionRange, formulaIds(itemIndex))
    Next itemIndex
    testDocument.Fields.Update

    regressionStage = "verify-seq-and-visible-ref"
    For itemIndex = 1 To formulaCount
        Set formulaRange = VTNumberedFormulaRangeForId( _
            testDocument, formulaIds(itemIndex))
        If formulaRange Is Nothing Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "An embedded-SEQ image formula could not be resolved."
        End If
        Set sequenceField = VTEquationSequenceFieldForBookmark( _
            testDocument, _
            VTEquationSequenceNumberBookmarkName(formulaIds(itemIndex)))
        Set visibleRef = Nothing
        For Each candidateField In formulaRange.Paragraphs(1).Range.Fields
            If candidateField.Type = wdFieldRef Then
                If VTReferenceTargetBookmarkName(candidateField.Code.Text) = _
                   VTEquationSequenceNumberBookmarkName(formulaIds(itemIndex)) Then
                    Set visibleRef = candidateField
                    Exit For
                End If
            End If
        Next candidateField
        If sequenceField Is Nothing Or visibleRef Is Nothing Or _
           Trim$(sequenceField.Result.Text) <> CStr(itemIndex) Or _
           Trim$(visibleRef.Result.Text) <> CStr(itemIndex) Or _
           sequenceField.Result.Paragraphs(1).Range.Start <> _
               formulaRange.Paragraphs(1).Range.Start Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The embedded SEQ or visible REF identity is incomplete."
        End If
    Next itemIndex

    For Each candidateField In testDocument.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            sequenceCount = sequenceCount + 1
        End If
    Next candidateField
    If sequenceCount <> formulaCount Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The embedded image prototype created an unexpected SEQ inventory."
    End If

    regressionStage = "word-built-in-cross-reference-list"
    items = testDocument.GetCrossReferenceItems(wdCaptionEquation)
    If Not IsArray(items) Or UBound(items) < formulaCount Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not expose embedded image SEQs in its Equation list."
    End If

    regressionStage = "word-built-in-cross-reference-insert"
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            bodyFieldCountBefore = bodyFieldCountBefore + 1
        End If
    Next candidateField
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    insertionRange.Select
    Selection.InsertCrossReference _
        ReferenceType:=wdCaptionEquation, _
        ReferenceKind:=wdOnlyLabelAndNumber, _
        ReferenceItem:=2, _
        InsertAsHyperlink:=True, _
        IncludePosition:=False
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            bodyFieldCountAfter = bodyFieldCountAfter + 1
            If bodyFieldCountAfter > bodyFieldCountBefore Then
                Set builtInField = candidateField
            End If
        End If
    Next candidateField
    If bodyFieldCountAfter <> bodyFieldCountBefore + 1 Or _
       builtInField Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not insert one built-in cross-reference to embedded SEQ."
    End If
    builtInTargetName = VTReferenceTargetBookmarkName( _
        builtInField.Code.Text)
    builtInResult = VTComparableEquationNumberText( _
        builtInField.Result.Text)
    If builtInResult <> "2" Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word built-in cross-reference to embedded SEQ is not pure" & _
            " [target=" & builtInTargetName & _
            "; result=" & builtInField.Result.Text & "]."
    End If

    regressionStage = "return-after-embedded-seq-row"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaIds(2))
    Set insertionRange = testDocument.Range( _
        Start:=formulaRange.Paragraphs(1).Range.End - 1, _
        End:=formulaRange.Paragraphs(1).Range.End - 1)
    insertionRange.Select
    Selection.TypeParagraph
    DoEvents
    Set blankParagraph = Selection.Range.Paragraphs(1).Range.Duplicate
    If blankParagraph.Fields.Count <> 0 Or _
       blankParagraph.InlineShapes.Count <> 0 Or _
       blankParagraph.OMaths.Count <> 0 Or _
       VTCustomTabStopCount(blankParagraph) <> 0 Or _
       VTWordRangeHasMeaningfulText(blankParagraph) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Return after embedded SEQ image did not create a plain paragraph."
    End If
    testDocument.Fields.Update
    For itemIndex = 1 To formulaCount
        Set sequenceField = VTEquationSequenceFieldForBookmark( _
            testDocument, _
            VTEquationSequenceNumberBookmarkName(formulaIds(itemIndex)))
        If sequenceField Is Nothing Or _
           Trim$(sequenceField.Result.Text) <> CStr(itemIndex) Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "Return changed embedded image SEQ ordering."
        End If
    Next itemIndex

    regressionStage = "middle-insert-embedded-seq"
    formulaIds(3) = VTRegressionPerformanceFormulaId(940000004)
    Set formulaShape = VTRegressionCreateEmbeddedSeqImageAtRange( _
        testDocument, blankParagraph, formulaIds(3))
    testDocument.Fields.Update
    sequenceCount = 0
    For Each candidateField In testDocument.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            sequenceCount = sequenceCount + 1
        End If
    Next candidateField
    If sequenceCount <> 4 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Middle embedded image insertion did not create one new SEQ."
    End If

    resultText = "PASS" & vbLf & _
        "structure=<hidden SEQ><TAB><image><TAB>(REF)<CR>" & vbLf & _
        "externalHelperParagraphs=0" & vbLf & _
        "wordEquationList=PASS" & vbLf & _
        "wordBuiltInCrossReference=PASS" & vbLf & _
        "wordBuiltInTarget=" & builtInTargetName & vbLf & _
        "return=PASS" & vbLf & _
        "middleInsert=PASS" & vbLf
    VTWriteTextAtomic resultPath, resultText
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

PrototypeFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    resultText = "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    VTWriteTextAtomic resultPath, resultText
    If Not testDocument Is Nothing Then
        testDocument.Saved = True
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word embedded image SEQ prototype", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordPreseedPrivateRefImagePrototype()
    Const formulaCount As Long = 3

    Dim testDocument As Document
    Dim formulaIds(1 To formulaCount) As String
    Dim formulaShape As InlineShape
    Dim insertionRange As Range
    Dim sequenceField As Field
    Dim candidateField As Field
    Dim builtInField As Field
    Dim items As Variant
    Dim resultPath As String
    Dim resultText As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim privateBookmarkName As String
    Dim builtInTargetName As String
    Dim builtInResult As String
    Dim itemIndex As Long
    Dim bodyFieldCountBefore As Long
    Dim bodyFieldCountAfter As Long

    On Error GoTo PrototypeFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-preseed-private-ref-image-prototype-result.txt"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate

    regressionStage = "create-and-preseed-private-bookmarks"
    For itemIndex = 1 To formulaCount
        formulaIds(itemIndex) = _
            VTRegressionPerformanceFormulaId(945000000 + itemIndex)
        Set insertionRange = VTAppendRegressionParagraph(testDocument)
        Set formulaShape = VTRegressionCreateEmbeddedSeqImageAtRange( _
            testDocument, insertionRange, formulaIds(itemIndex))
        Set sequenceField = VTEquationSequenceFieldForBookmark( _
            testDocument, _
            VTEquationSequenceNumberBookmarkName(formulaIds(itemIndex)))
        If sequenceField Is Nothing Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The preseed prototype could not resolve its SEQ field."
        End If
        privateBookmarkName = "_Ref945000" & Right$("000" & CStr(itemIndex), 3)
        If testDocument.Bookmarks.Exists(privateBookmarkName) Then _
            testDocument.Bookmarks(privateBookmarkName).Delete
        testDocument.Bookmarks.Add _
            Name:=privateBookmarkName, Range:=sequenceField.Result.Duplicate
    Next itemIndex
    testDocument.Fields.Update

    regressionStage = "verify-equation-list"
    items = testDocument.GetCrossReferenceItems(wdCaptionEquation)
    If Not IsArray(items) Or UBound(items) < formulaCount Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not expose preseeded image SEQs in its Equation list."
    End If

    regressionStage = "insert-built-in-reference"
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            bodyFieldCountBefore = bodyFieldCountBefore + 1
        End If
    Next candidateField
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    insertionRange.Select
    Selection.InsertCrossReference _
        ReferenceType:=wdCaptionEquation, _
        ReferenceKind:=wdOnlyLabelAndNumber, _
        ReferenceItem:=2, _
        InsertAsHyperlink:=True, _
        IncludePosition:=False
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            bodyFieldCountAfter = bodyFieldCountAfter + 1
            If bodyFieldCountAfter > bodyFieldCountBefore Then
                Set builtInField = candidateField
            End If
        End If
    Next candidateField
    If bodyFieldCountAfter <> bodyFieldCountBefore + 1 Or _
       builtInField Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not insert one built-in cross-reference in the preseed prototype."
    End If
    builtInTargetName = VTReferenceTargetBookmarkName( _
        builtInField.Code.Text)
    builtInResult = VTComparableEquationNumberText( _
        builtInField.Result.Text)
    If builtInResult <> "2" Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Preseeding a private _Ref did not keep Word's result pure" & _
            " [target=" & builtInTargetName & _
            "; result=" & builtInField.Result.Text & "]."
    End If

    resultText = "PASS" & vbLf & _
        "preseedPrivateRef=PASS" & vbLf & _
        "builtInTarget=" & builtInTargetName & vbLf & _
        "builtInResult=" & builtInField.Result.Text & vbLf
    VTWriteTextAtomic resultPath, resultText
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

PrototypeFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    resultText = "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    VTWriteTextAtomic resultPath, resultText
    If Not testDocument Is Nothing Then
        testDocument.Saved = True
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word preseed private-ref prototype", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Private Function VTRegressionCreateLeadingHelperImageAtRange( _
    ByVal documentObject As Document, _
    ByVal requestedRange As Range, _
    ByVal formulaId As String) As InlineShape

    Dim helperParagraph As Range
    Dim formulaParagraph As Range
    Dim formulaInsertion As Range
    Dim formulaShape As InlineShape
    Dim formulaContainerRange As Range
    Dim insertionRange As Range
    Dim fieldRange As Range
    Dim sequenceField As Field
    Dim visibleRef As Field
    Dim numberRange As Range
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim captionBookmarkName As String
    Dim helperStart As Long
    Dim formulaStart As Long

    If documentObject Is Nothing Or requestedRange Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The leading-helper image prototype target is invalid."
    End If

    Set helperParagraph = requestedRange.Paragraphs(1).Range.Duplicate
    VTNormalizePlainWordParagraph helperParagraph
    helperStart = helperParagraph.Start
    Set fieldRange = documentObject.Range( _
        Start:=helperStart, End:=helperStart)
    Set sequenceField = VTInsertRegisteredEquationCaption( _
        fieldRange, VTNativeEquationLabelName())
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    VTFormatHiddenEquationParagraph helperParagraph

    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then _
        documentObject.Bookmarks(sequenceBookmarkName).Delete
    documentObject.Bookmarks.Add _
        Name:=sequenceBookmarkName, Range:=sequenceField.Result.Duplicate
    VTSetCollapsedEquationCaptionBookmark _
        documentObject, formulaId, helperParagraph

    ' The visible formula paragraph follows its sequence source. Return at the
    ' end of the formula can only create content after the formula; it can never
    ' separate this formula from the SEQ that already precedes it.
    helperParagraph.InsertParagraphAfter
    Set helperParagraph = documentObject.Range( _
        Start:=helperStart, End:=helperStart).Paragraphs(1).Range.Duplicate
    formulaStart = helperParagraph.End
    Set formulaInsertion = documentObject.Range( _
        Start:=formulaStart, End:=formulaStart)
    Set formulaShape = documentObject.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=formulaInsertion)
    formulaShape.Width = 90!
    formulaShape.Height = 24!
    formulaShape.AlternativeText = VT_METADATA_PREFIX & "e30"
    formulaShape.Title = VTFormulaReference(formulaId, "block", True)

    Set formulaParagraph = formulaShape.Range.Paragraphs(1).Range.Duplicate
    VTConfigureNumberedEquationParagraph formulaParagraph
    Set insertionRange = documentObject.Range( _
        Start:=formulaShape.Range.Start, End:=formulaShape.Range.Start)
    insertionRange.Text = vbTab
    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, formulaStart)
    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set insertionRange = documentObject.Range( _
        Start:=formulaContainerRange.End, End:=formulaContainerRange.End)
    insertionRange.Text = vbTab & "()"

    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, formulaStart)
    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set fieldRange = documentObject.Range( _
        Start:=formulaContainerRange.End + 2, _
        End:=formulaContainerRange.End + 2)
    Set visibleRef = fieldRange.Fields.Add( _
        Range:=fieldRange, Type:=wdFieldRef, _
        Text:=sequenceBookmarkName, PreserveFormatting:=False)
    visibleRef.Update
    With visibleRef.Result.Font
        .Hidden = False
        .Color = wdColorAutomatic
        .Name = VT_WORD_EQUATION_NUMBER_FONT_NAME
        .NameAscii = VT_WORD_EQUATION_NUMBER_FONT_NAME
        .NameOther = VT_WORD_EQUATION_NUMBER_FONT_NAME
        .Position = 0
        .Size = VTVisibleEquationNumberFontSize(documentObject)
    End With

    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, formulaStart)
    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set formulaParagraph = formulaShape.Range.Paragraphs(1).Range.Duplicate
    Set numberRange = documentObject.Range( _
        Start:=formulaContainerRange.End + 1, _
        End:=formulaParagraph.End - 1)
    If numberRange.Fields.Count <> 1 Or _
       Left$(numberRange.Text, 1) <> "(" Or _
       Right$(numberRange.Text, 1) <> ")" Or _
       VTReferenceTargetBookmarkName( _
           numberRange.Fields(1).Code.Text) <> sequenceBookmarkName Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not keep the visible REF beside the leading-helper image."
    End If
    If documentObject.Bookmarks.Exists(numberBookmarkName) Then _
        documentObject.Bookmarks(numberBookmarkName).Delete
    documentObject.Bookmarks.Add _
        Name:=numberBookmarkName, Range:=numberRange

    Set VTRegressionCreateLeadingHelperImageAtRange = formulaShape
End Function

Public Sub VisualTeX_RunWordLeadingHelperImagePrototype()
    Const formulaCount As Long = 3

    Dim testDocument As Document
    Dim formulaIds(1 To formulaCount) As String
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim formulaParagraph As Range
    Dim helperParagraph As Range
    Dim insertionRange As Range
    Dim blankParagraph As Range
    Dim sequenceField As Field
    Dim candidateField As Field
    Dim builtInField As Field
    Dim visibleRef As Field
    Dim items As Variant
    Dim resultPath As String
    Dim resultText As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim builtInTargetName As String
    Dim builtInResult As String
    Dim itemIndex As Long
    Dim bodyFieldCountBefore As Long
    Dim bodyFieldCountAfter As Long
    Dim helperStart As Long

    On Error GoTo PrototypeFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-leading-helper-image-prototype-result.txt"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate

    regressionStage = "create-three-leading-helper-images"
    For itemIndex = 1 To formulaCount
        formulaIds(itemIndex) = _
            VTRegressionPerformanceFormulaId(950000000 + itemIndex)
        Set insertionRange = VTAppendRegressionParagraph(testDocument)
        Set formulaShape = VTRegressionCreateLeadingHelperImageAtRange( _
            testDocument, insertionRange, formulaIds(itemIndex))
    Next itemIndex
    testDocument.Fields.Update

    regressionStage = "verify-helper-before-formula"
    For itemIndex = 1 To formulaCount
        Set formulaRange = VTNumberedFormulaRangeForId( _
            testDocument, formulaIds(itemIndex))
        Set sequenceField = VTEquationSequenceFieldForBookmark( _
            testDocument, _
            VTEquationSequenceNumberBookmarkName(formulaIds(itemIndex)))
        If formulaRange Is Nothing Or sequenceField Is Nothing Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "A leading-helper formula identity is missing."
        End If
        Set formulaParagraph = formulaRange.Paragraphs(1).Range.Duplicate
        Set helperParagraph = _
            sequenceField.Result.Paragraphs(1).Range.Duplicate
        If Not VTHelperParagraphOwnsNativeEquationSequence( _
               helperParagraph) Or _
           helperParagraph.End <> formulaParagraph.Start Or _
           Trim$(sequenceField.Result.Text) <> CStr(itemIndex) Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The Equation helper is not immediately before its image formula" & _
                " [helper=" & CStr(helperParagraph.Start) & "-" & _
                    CStr(helperParagraph.End) & _
                "; formula=" & CStr(formulaParagraph.Start) & "-" & _
                    CStr(formulaParagraph.End) & "]."
        End If
        Set visibleRef = Nothing
        For Each candidateField In formulaParagraph.Fields
            If candidateField.Type = wdFieldRef And _
               VTReferenceTargetBookmarkName(candidateField.Code.Text) = _
                   VTEquationSequenceNumberBookmarkName(formulaIds(itemIndex)) Then
                Set visibleRef = candidateField
                Exit For
            End If
        Next candidateField
        If visibleRef Is Nothing Or _
           Trim$(visibleRef.Result.Text) <> CStr(itemIndex) Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The visible image number REF is not synchronized to its leading helper."
        End If
    Next itemIndex

    regressionStage = "word-built-in-cross-reference-list"
    items = testDocument.GetCrossReferenceItems(wdCaptionEquation)
    If Not IsArray(items) Or UBound(items) < formulaCount Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not expose leading helpers in its Equation list."
    End If

    regressionStage = "word-built-in-cross-reference-insert"
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            bodyFieldCountBefore = bodyFieldCountBefore + 1
        End If
    Next candidateField
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    insertionRange.Select
    Selection.InsertCrossReference _
        ReferenceType:=wdCaptionEquation, _
        ReferenceKind:=wdOnlyLabelAndNumber, _
        ReferenceItem:=2, _
        InsertAsHyperlink:=True, _
        IncludePosition:=False
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            bodyFieldCountAfter = bodyFieldCountAfter + 1
            If bodyFieldCountAfter > bodyFieldCountBefore Then
                Set builtInField = candidateField
            End If
        End If
    Next candidateField
    If bodyFieldCountAfter <> bodyFieldCountBefore + 1 Or _
       builtInField Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not insert one built-in cross-reference to leading helper."
    End If
    builtInTargetName = VTReferenceTargetBookmarkName( _
        builtInField.Code.Text)
    builtInResult = VTComparableEquationNumberText( _
        builtInField.Result.Text)
    If builtInResult <> "2" Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word built-in cross-reference to leading helper is not pure" & _
            " [target=" & builtInTargetName & _
            "; result=" & builtInField.Result.Text & "]."
    End If

    regressionStage = "return-after-formula"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaIds(2))
    Set formulaParagraph = formulaRange.Paragraphs(1).Range.Duplicate
    helperStart = VTEquationSequenceFieldForBookmark( _
        testDocument, _
        VTEquationSequenceNumberBookmarkName(formulaIds(2))). _
            Result.Paragraphs(1).Range.Start
    Set insertionRange = testDocument.Range( _
        Start:=formulaParagraph.End - 1, End:=formulaParagraph.End - 1)
    insertionRange.Select
    Selection.TypeParagraph
    DoEvents
    Set blankParagraph = Selection.Range.Paragraphs(1).Range.Duplicate
    Set sequenceField = VTEquationSequenceFieldForBookmark( _
        testDocument, _
        VTEquationSequenceNumberBookmarkName(formulaIds(2)))
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaIds(2))
    Set formulaParagraph = formulaRange.Paragraphs(1).Range.Duplicate
    If helperParagraph.Start <> helperStart Or _
       helperParagraph.End <> formulaParagraph.Start Or _
       blankParagraph.Start < formulaParagraph.End Or _
       blankParagraph.Fields.Count <> 0 Or _
       blankParagraph.InlineShapes.Count <> 0 Or _
       blankParagraph.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(blankParagraph) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Return displaced a leading Equation helper from its image formula."
    End If

    regressionStage = "middle-insert-leading-helper"
    formulaIds(3) = VTRegressionPerformanceFormulaId(950000004)
    Set formulaShape = VTRegressionCreateLeadingHelperImageAtRange( _
        testDocument, blankParagraph, formulaIds(3))
    testDocument.Fields.Update
    Set sequenceField = VTEquationSequenceFieldForBookmark( _
        testDocument, _
        VTEquationSequenceNumberBookmarkName(formulaIds(3)))
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaIds(3))
    If sequenceField Is Nothing Or formulaRange Is Nothing Or _
       sequenceField.Result.Paragraphs(1).Range.End <> _
           formulaRange.Paragraphs(1).Range.Start Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The inserted leading helper is not attached to its new formula."
    End If

    resultText = "PASS" & vbLf & _
        "structure=<helper SEQ><CR><TAB><image><TAB>(REF)<CR>" & vbLf & _
        "helperBeforeFormula=PASS" & vbLf & _
        "wordEquationList=PASS" & vbLf & _
        "wordBuiltInCrossReference=PASS" & vbLf & _
        "wordBuiltInTarget=" & builtInTargetName & vbLf & _
        "returnKeepsHelperAttached=PASS" & vbLf & _
        "middleInsert=PASS" & vbLf
    VTWriteTextAtomic resultPath, resultText
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

PrototypeFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    resultText = "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    VTWriteTextAtomic resultPath, resultText
    If Not testDocument Is Nothing Then
        testDocument.Saved = True
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word leading-helper image prototype", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Private Function VTRegressionAppendChapterHeading( _
    ByVal documentObject As Document, _
    ByVal headingText As String, _
    ByVal useBuiltInHeadingStyle As Boolean) As Range

    Dim insertionRange As Range
    Dim headingRange As Range
    Dim headingStart As Long

    If documentObject Is Nothing Or Len(headingText) = 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The chapter-numbering regression heading target is invalid."
    End If
    Set insertionRange = VTAppendRegressionParagraph(documentObject)
    headingStart = insertionRange.Start
    insertionRange.Text = headingText
    Set headingRange = documentObject.Range( _
        Start:=headingStart, End:=headingStart + Len(headingText))
    Set headingRange = headingRange.Paragraphs(1).Range.Duplicate
    If useBuiltInHeadingStyle Then
        headingRange.Style = wdStyleHeading1
    Else
        headingRange.Style = wdStyleNormal
        headingRange.ParagraphFormat.OutlineLevel = wdOutlineLevel1
    End If
    If headingRange.Paragraphs(1).OutlineLevel <> wdOutlineLevel1 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The chapter-numbering regression heading did not retain level 1."
    End If
    Set VTRegressionAppendChapterHeading = headingRange.Duplicate
End Function

Private Function VTRegressionAppendSectionHeading( _
    ByVal documentObject As Document, _
    ByVal headingText As String, _
    ByVal useBuiltInHeadingStyle As Boolean) As Range

    Dim insertionRange As Range
    Dim headingRange As Range
    Dim headingStart As Long

    If documentObject Is Nothing Or Len(headingText) = 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The section-numbering regression heading target is invalid."
    End If
    Set insertionRange = VTAppendRegressionParagraph(documentObject)
    headingStart = insertionRange.Start
    insertionRange.Text = headingText
    Set headingRange = documentObject.Range( _
        Start:=headingStart, End:=headingStart + Len(headingText))
    Set headingRange = headingRange.Paragraphs(1).Range.Duplicate
    If useBuiltInHeadingStyle Then
        headingRange.Style = wdStyleHeading2
    Else
        headingRange.Style = wdStyleNormal
        headingRange.ParagraphFormat.OutlineLevel = wdOutlineLevel2
    End If
    If headingRange.Paragraphs(1).OutlineLevel <> wdOutlineLevel2 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The section-numbering regression heading did not retain level 2."
    End If
    Set VTRegressionAppendSectionHeading = headingRange.Duplicate
End Function

Private Sub VTRegressionAssertChapterNumberOrdinal( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal expectedOrdinal As Long, _
    Optional ByVal expectedPrefix As String = "")

    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim sequenceField As Field
    Dim placeRef As Field
    Dim numberRange As Range
    Dim formulaShape As InlineShape
    Dim candidateId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim imageFormula As Boolean
    Dim expectedSuffix As String
    Dim expectedNumberText As String
    Dim restartLevel As Long

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Or _
       expectedOrdinal < 1 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The chapter-numbering regression assertion target is invalid."
    End If
    sequenceBookmarkName = VTEquationSequenceNumberBookmarkName(formulaId)
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    For Each formulaShape In documentObject.InlineShapes
        If VTTryParseFormulaReference( _
           formulaShape.Title, candidateId, displayMode, numbered) Then
            If StrComp(candidateId, formulaId, vbTextCompare) = 0 Then
                Set placeRef = VTImagePlaceRef(formulaShape.Range)
                If Not placeRef Is Nothing Then
                    imageFormula = True
                    Set sequenceField = VTImageSequenceField(placeRef, False)
                End If
                Exit For
            End If
        End If
    Next formulaShape
    If Not imageFormula Then
        Set sequenceField = VTEquationSequenceFieldForBookmark( _
            documentObject, sequenceBookmarkName)
    End If
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The chapter-numbering regression lost its SEQ field."
    End If
    restartLevel = VTEquationNumberingRestartLevel(documentObject)
    If imageFormula Then
        If StrComp(VTSequenceIdentifierFromFieldCode(sequenceField.Code.Text), _
           VT_WORD_IMAGE_SEQUENCE_NAME, vbTextCompare) <> 0 Or _
           VTEquationSequenceResultText(sequenceField) <> _
               CStr(expectedOrdinal) Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The private image SEQ result is stale" & _
                " [code=" & sequenceField.Code.Text & _
                "; actual=" & VTEquationSequenceResultText(sequenceField) & _
                "; expected=" & CStr(expectedOrdinal) & "]."
        End If
    ElseIf Not VTEquationSequenceFieldHasOrdinal( _
       sequenceField, VTNativeEquationLabelName(), _
       expectedOrdinal, restartLevel) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The chapter-numbering regression SEQ result is stale" & _
            " [code=" & sequenceField.Code.Text & _
            "; actual=" & VTEquationSequenceResultText(sequenceField) & _
            "; expected=" & CStr(expectedOrdinal) & "]."
    End If
    If Not documentObject.Bookmarks.Exists(numberBookmarkName) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The chapter-numbering regression lost its visible number Bookmark."
    End If
    Set numberRange = documentObject.Bookmarks( _
        numberBookmarkName).Range.Duplicate
    If Len(expectedPrefix) > 0 Then
        expectedNumberText = "(" & expectedPrefix & _
            VTEquationNumberingDisplaySeparator(documentObject) & _
            CStr(expectedOrdinal) & ")"
        If numberRange.Text <> expectedNumberText Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The chapter-numbering regression visible number is stale" & _
                " [actual=" & numberRange.Text & _
                "; expected=" & expectedNumberText & "]."
        End If
    Else
        expectedSuffix = CStr(expectedOrdinal) & ")"
        If Len(numberRange.Text) < Len(expectedSuffix) Or _
           Right$(numberRange.Text, Len(expectedSuffix)) <> expectedSuffix Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The chapter-numbering regression visible number is stale" & _
                " [actual=" & numberRange.Text & _
                "; expectedSuffix=" & expectedSuffix & "]."
        End If
    End If
End Sub

Public Sub VisualTeX_RunWordChapterPrependNumberingRegression()
    Const latexBase64 As String = "eF4yK3leMg"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim insertionRange As Range
    Dim insertedReference As Range
    Dim builtInTargetRange As Range
    Dim insertedFormulaParagraph As Range
    Dim insertedHelperParagraph As Range
    Dim insertedNumberRange As Range
    Dim sequenceField As Field
    Dim candidateField As Field
    Dim insertedVisibleReferenceField As Field
    Dim capturedBindings As Collection
    Dim bindingItem As Variant
    Dim resultPath As String
    Dim resultText As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim firstImageId As String
    Dim firstNativeId As String
    Dim secondImageId As String
    Dim secondNativeId As String
    Dim insertedImageId As String
    Dim referenceTargetName As String
    Dim referenceResult As String
    Dim builtInTargetName As String
    Dim builtInResult As String
    Dim bodyFieldCountBefore As Long
    Dim bodyFieldCountAfter As Long
    Dim newestBodyFieldStart As Long
    Dim candidateFieldStart As Long
    Dim expectedVisibleSize As Single
    Dim builtInBindingCaptured As Boolean
    Dim visibleReferenceFormattingValid As Boolean
    Dim numberRangeFormattingValid As Boolean
    Dim helperFormattingValid As Boolean

    On Error GoTo RegressionFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-chapter-prepend-numbering-regression-result.txt"

    firstImageId = VTRegressionPerformanceFormulaId(910000001)
    firstNativeId = VTRegressionPerformanceFormulaId(910000002)
    secondImageId = VTRegressionPerformanceFormulaId(910000003)
    secondNativeId = VTRegressionPerformanceFormulaId(910000004)
    insertedImageId = VTRegressionPerformanceFormulaId(919999999)
    referenceTargetName = _
        VTEquationSequenceNumberBookmarkName(secondNativeId)

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_CHAPTER, "-"

    regressionStage = "chapter-1-standard-heading-1"
    VTRegressionAppendChapterHeading testDocument, "Chapter One", True
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, firstImageId, latexBase64, 96!, 28!)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, firstNativeId, latexBase64, 104!, 30!)

    regressionStage = "chapter-2-custom-outline-level"
    VTRegressionAppendChapterHeading testDocument, "Chapter Two", False
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, secondImageId, latexBase64, 96!, 28!)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, secondNativeId, latexBase64, 104!, 30!)

    regressionStage = "verify-two-real-chapters"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, firstImageId, 1, "1"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, firstNativeId, 2, "1"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, secondImageId, 1, "2"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, secondNativeId, 2, "2"

    regressionStage = "create-body-cross-reference"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set insertedReference = VTInsertEquationNumberReferenceAtRange( _
        insertionRange, 4, True)
    If insertedReference.Text <> "(" & _
       VTEquationNumberTextForFormula(testDocument, secondNativeId) & ")" Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The chapter-numbering regression body REF was not inserted."
    End If

    regressionStage = "insert-inside-second-chapter"
    Set insertionRange = VTRegressionInsertBlankParagraphBeforeFormula( _
        testDocument, secondNativeId)
    Set formulaShape = VTRegressionCreateNumberedImageAtCaret( _
        testDocument, insertionRange, insertedImageId, _
        latexBase64, 96!, 28!)
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, firstImageId, 1, "1"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, firstNativeId, 2, "1"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, secondImageId, 1, "2"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, insertedImageId, 2, "2"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, secondNativeId, 3, "2"
    referenceResult = VTBodyReferenceResultForTarget( _
        testDocument, referenceTargetName)
    If referenceResult <> VTEquationNumberTextForFormula( _
       testDocument, secondNativeId) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The body REF lost its target after second-chapter insertion."
    End If

    regressionStage = "create-word-built-in-cross-reference"
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            bodyFieldCountBefore = bodyFieldCountBefore + 1
        End If
    Next candidateField
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    insertionRange.Select
    Selection.InsertCrossReference _
        ReferenceType:=wdCaptionEquation, _
        ReferenceKind:=wdOnlyLabelAndNumber, _
        ReferenceItem:=4, _
        InsertAsHyperlink:=True, _
        IncludePosition:=False
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            bodyFieldCountAfter = bodyFieldCountAfter + 1
            candidateFieldStart = VTEquationFieldStart(candidateField)
            If candidateFieldStart > newestBodyFieldStart Then
                newestBodyFieldStart = candidateFieldStart
                builtInTargetName = VTReferenceTargetBookmarkName( _
                    candidateField.Code.Text)
            End If
        End If
    Next candidateField
    If Len(builtInTargetName) = 0 Or _
       bodyFieldCountAfter <> bodyFieldCountBefore + 1 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not create the chapter regression built-in Equation REF."
    End If

    regressionStage = "verify-word-built-in-private-ref-span"
    If Not testDocument.Bookmarks.Exists(builtInTargetName) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not retain the built-in private Equation Bookmark."
    End If
    Set builtInTargetRange = _
        testDocument.Bookmarks(builtInTargetName).Range.Duplicate
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, insertedImageId)
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, insertedImageId)
    If formulaRange Is Nothing Or sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The built-in private Equation Bookmark lost its formula/helper target."
    End If
    Set insertedFormulaParagraph = _
        formulaRange.Paragraphs(1).Range.Duplicate
    Set insertedHelperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If insertedHelperParagraph.Start <> insertedFormulaParagraph.End Or _
       builtInTargetRange.Start < insertedHelperParagraph.Start Or _
       builtInTargetRange.End > insertedHelperParagraph.End - 1 Or _
       builtInTargetRange.InlineShapes.Count <> 0 Or _
       builtInTargetRange.OMaths.Count <> 0 Or _
       InStr(1, builtInTargetRange.Text, vbTab, vbBinaryCompare) > 0 Or _
       InStr(1, builtInTargetRange.Text, vbCr, vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word built-in _Ref escaped the external SEQ helper" & _
            " [formula=" & CStr(insertedFormulaParagraph.Start) & "-" & _
                CStr(insertedFormulaParagraph.End) & _
            "; helper=" & CStr(insertedHelperParagraph.Start) & "-" & _
                CStr(insertedHelperParagraph.End) & _
            "; target=" & CStr(builtInTargetRange.Start) & "-" & _
                CStr(builtInTargetRange.End) & _
            "; text=" & Replace$(Replace$(builtInTargetRange.Text, _
                vbTab, "<TAB>"), vbCr, "<CR>") & _
            "; images=" & CStr(builtInTargetRange.InlineShapes.Count) & "]."
    End If
    If VTComparableEquationNumberText(builtInTargetRange.Text) <> _
       VTComparableEquationNumberText( _
           VTEquationNumberTextForFormula(testDocument, insertedImageId)) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word built-in _Ref is not the helper's pure formatted number" & _
            " [target=" & builtInTargetRange.Text & _
            "; expected=" & VTEquationNumberTextForFormula( _
                testDocument, insertedImageId) & "]."
    End If

    regressionStage = "f9-visible-ref-font"
    testDocument.Fields.Update
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, insertedImageId)
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "F9 lost the visible image Equation."
    End If
    Set insertedVisibleReferenceField = VTImageEquationReferenceField( _
        formulaRange, insertedImageId)
    If Not insertedVisibleReferenceField Is Nothing Then
        Set insertedNumberRange = VTImageEquationNumberRange( _
            formulaRange, insertedVisibleReferenceField)
    End If
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, insertedImageId)
    expectedVisibleSize = VTVisibleEquationNumberFontSize(testDocument)
    If insertedVisibleReferenceField Is Nothing Or _
       insertedNumberRange Is Nothing Or sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "F9 lost the image Equation's external SEQ or visible REF."
    End If
    visibleReferenceFormattingValid = _
        StrComp(CStr(insertedVisibleReferenceField.Result.Font.NameAscii), _
            VT_WORD_EQUATION_NUMBER_FONT_NAME, vbTextCompare) = 0 And _
        StrComp(CStr(insertedVisibleReferenceField.Result.Font.NameOther), _
            VT_WORD_EQUATION_NUMBER_FONT_NAME, vbTextCompare) = 0 And _
        insertedVisibleReferenceField.Result.Font.Position = 0
    numberRangeFormattingValid = _
        insertedNumberRange.Font.Position = 0 And _
        insertedNumberRange.Font.Color = wdColorAutomatic And _
        StrComp(CStr(insertedNumberRange.Font.NameAscii), _
            VT_WORD_EQUATION_NUMBER_FONT_NAME, vbTextCompare) = 0 And _
        StrComp(CStr(insertedNumberRange.Font.NameOther), _
            VT_WORD_EQUATION_NUMBER_FONT_NAME, vbTextCompare) = 0 And _
        Abs(insertedNumberRange.Font.Size - expectedVisibleSize) <= 0.1
    helperFormattingValid = _
        StrComp(CStr(sequenceField.Result.Font.NameAscii), _
            VT_WORD_EQUATION_NUMBER_FONT_NAME, vbTextCompare) = 0 And _
        StrComp(CStr(sequenceField.Result.Font.NameOther), _
            VT_WORD_EQUATION_NUMBER_FONT_NAME, vbTextCompare) = 0 And _
        sequenceField.Result.Font.Position = 0 And _
        sequenceField.Result.Font.Color = wdColorAutomatic And _
        Abs(sequenceField.Result.Font.Size - expectedVisibleSize) <= 0.1
    If insertedNumberRange.Text <> "(" & _
           VTEquationNumberTextForFormula(testDocument, insertedImageId) & ")" Or _
       Not visibleReferenceFormattingValid Or _
       Not numberRangeFormattingValid Or Not helperFormattingValid Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "F9 changed the image Equation number away from Cambria Math Position 0" & _
            " [visibleAscii=" & CStr( _
                insertedVisibleReferenceField.Result.Font.NameAscii) & _
            "; visibleOther=" & CStr( _
                insertedVisibleReferenceField.Result.Font.NameOther) & _
            "; visiblePosition=" & CStr( _
                insertedVisibleReferenceField.Result.Font.Position) & _
            "; numberAscii=" & CStr(insertedNumberRange.Font.NameAscii) & _
            "; helperAscii=" & CStr(sequenceField.Result.Font.NameAscii) & _
            "; helperOther=" & CStr( _
                sequenceField.Result.Font.NameOther) & _
            "; helperPosition=" & CStr( _
                sequenceField.Result.Font.Position) & "]."
    End If

    Set capturedBindings = _
        VTCaptureBodyEquationReferenceBindings(testDocument)
    builtInBindingCaptured = False
    For Each bindingItem In capturedBindings
        If InStr(1, CStr(bindingItem), _
           builtInTargetName & vbTab & insertedImageId & vbTab, _
           vbTextCompare) = 1 Then
            builtInBindingCaptured = True
            Exit For
        End If
    Next bindingItem
    If Not builtInBindingCaptured Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word built-in _Ref was not bound to the inserted chapter formula."
    End If

    regressionStage = "delete-first-chapter-formula-and-update"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, firstImageId)
    If formulaRange Is Nothing Or formulaRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The first chapter image formula could not be deleted."
    End If
    formulaRange.InlineShapes(1).Delete
    VTPruneOrphanedEquationNumberScaffolds testDocument
    VTUpdateEquationNumbersCore testDocument, True
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, firstNativeId, 1, "1"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, secondImageId, 1, "2"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, insertedImageId, 2, "2"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, secondNativeId, 3, "2"

    regressionStage = "delete-second-chapter-formula-and-update"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, secondImageId)
    If formulaRange Is Nothing Or formulaRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The second chapter image formula could not be deleted."
    End If
    formulaRange.InlineShapes(1).Delete
    VTPruneOrphanedEquationNumberScaffolds testDocument
    VTUpdateEquationNumbersCore testDocument, True
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, insertedImageId, 1, "2"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, secondNativeId, 2, "2"
    referenceResult = VTBodyReferenceResultForTarget( _
        testDocument, referenceTargetName)
    If referenceResult <> VTEquationNumberTextForFormula( _
       testDocument, secondNativeId) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The body REF lost its target after chapter deletion/update."
    End If
    builtInResult = VTBodyReferenceResultForTarget( _
        testDocument, builtInTargetName)
    If VTComparableEquationNumberText(builtInResult) <> _
       VTComparableEquationNumberText( _
           VTEquationNumberTextForFormula( _
               testDocument, insertedImageId)) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word built-in _Ref did not follow 2-2 to 2-1 after reset" & _
            " [result=" & builtInResult & "]."
    End If
    Set capturedBindings = _
        VTCaptureBodyEquationReferenceBindings(testDocument)
    builtInBindingCaptured = False
    For Each bindingItem In capturedBindings
        If InStr(1, CStr(bindingItem), _
           builtInTargetName & vbTab & insertedImageId & vbTab, _
           vbTextCompare) = 1 Then
            builtInBindingCaptured = True
            Exit For
        End If
    Next bindingItem
    If Not builtInBindingCaptured Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word built-in _Ref changed target when SEQ gained a \\r 1 reset."
    End If

    regressionStage = "explicit-update-repairs-stale-local-seq"
    Set sequenceField = VTEquationSequenceFieldForBookmark( _
        testDocument, referenceTargetName)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The chapter-numbering regression could not poison the local SEQ."
    End If
    sequenceField.Result.Text = "99"
    VTUpdateEquationNumbersCore testDocument, True
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, secondNativeId, 2, "2"
    referenceResult = VTBodyReferenceResultForTarget( _
        testDocument, referenceTargetName)
    If referenceResult <> VTEquationNumberTextForFormula( _
       testDocument, secondNativeId) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The body REF did not recover with the repaired local SEQ."
    End If

    resultText = "PASS" & vbLf & _
        "initial=1-1,1-2,2-1,2-2" & vbLf & _
        "standardHeading1=PASS" & vbLf & _
        "customOutlineLevel1=PASS" & vbLf & _
        "secondChapterInsertion=PASS" & vbLf & _
        "deleteAndUpdate=PASS" & vbLf & _
        "bodyCrossReference=PASS" & vbLf & _
        "wordBuiltInPrivateRef=PASS" & vbLf & _
        "privateRefSpan=helper-only" & vbLf & _
        "f9VisibleNumberFont=Cambria Math,Position 0" & vbLf & _
        "explicitUpdateRepair=PASS" & vbLf
    VTWriteTextAtomic resultPath, resultText
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    resultText = "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    VTWriteTextAtomic resultPath, resultText
    ' Preserve the failed Word document and selection for exact popup/process/
    ' range inspection. The external harness closes it without saving only after
    ' recording the failure evidence.
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word chapter prepend numbering regression", _
        regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordChapterZeroScopeNumberingRegression()
    Const latexBase64 As String = "eF4yK3leMg"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim resultPath As String
    Dim resultText As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim firstFormulaId As String
    Dim secondFormulaId As String
    Dim thirdFormulaId As String

    On Error GoTo RegressionFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-chapter-zero-scope-numbering-regression-result.txt"
    firstFormulaId = VTRegressionPerformanceFormulaId(911100001)
    secondFormulaId = VTRegressionPerformanceFormulaId(911100002)
    thirdFormulaId = VTRegressionPerformanceFormulaId(911100003)

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_CHAPTER, "-"

    regressionStage = "before-first-chapter"
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, firstFormulaId, latexBase64, 96!, 28!)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, secondFormulaId, latexBase64, 96!, 28!)
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, firstFormulaId, 1, "0"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, secondFormulaId, 2, "0"

    regressionStage = "enter-first-chapter"
    VTRegressionAppendChapterHeading testDocument, "Chapter One", True
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, thirdFormulaId, latexBase64, 96!, 28!)
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, thirdFormulaId, 1, "1"

    regressionStage = "explicit-update-retains-zero-chapter"
    VTUpdateEquationNumbersCore testDocument, True
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, firstFormulaId, 1, "0"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, secondFormulaId, 2, "0"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, thirdFormulaId, 1, "1"

    resultText = "PASS" & vbLf & _
        "initial=0-1,0-2,1-1" & vbLf & _
        "zeroChapterScope=PASS" & vbLf & _
        "explicitUpdateRepair=PASS" & vbLf
    VTWriteTextAtomic resultPath, resultText
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    resultText = "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    VTWriteTextAtomic resultPath, resultText
    ' Preserve the failed document for popup/process/range inspection.
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word chapter zero-scope numbering regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordSectionImplicitNumberingRegression()
    Const latexBase64 As String = "eF4yK3leMg"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim resultPath As String
    Dim resultText As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim preChapterFormulaId As String
    Dim chapterOneFirstId As String
    Dim chapterOneSecondId As String
    Dim chapterTwoFirstId As String
    Dim chapterTwoSecondId As String
    Dim sectionOneFirstId As String
    Dim sectionOneSecondId As String

    On Error GoTo RegressionFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-section-implicit-numbering-regression-result.txt"
    preChapterFormulaId = VTRegressionPerformanceFormulaId(911000000)
    chapterOneFirstId = VTRegressionPerformanceFormulaId(911000001)
    chapterOneSecondId = VTRegressionPerformanceFormulaId(911000002)
    chapterTwoFirstId = VTRegressionPerformanceFormulaId(911000003)
    chapterTwoSecondId = VTRegressionPerformanceFormulaId(911000004)
    sectionOneFirstId = VTRegressionPerformanceFormulaId(911000005)
    sectionOneSecondId = VTRegressionPerformanceFormulaId(911000006)

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SECTION, "."

    regressionStage = "before-any-chapter-or-section"
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, preChapterFormulaId, latexBase64, 96!, 28!)
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, preChapterFormulaId, 1, "0.0"

    regressionStage = "chapter-1-before-first-section"
    VTRegressionAppendChapterHeading testDocument, "Chapter One", True
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, chapterOneFirstId, latexBase64, 96!, 28!)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, chapterOneSecondId, latexBase64, 96!, 28!)

    regressionStage = "chapter-2-before-first-section"
    VTRegressionAppendChapterHeading testDocument, "Chapter Two", False
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, chapterTwoFirstId, latexBase64, 96!, 28!)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, chapterTwoSecondId, latexBase64, 96!, 28!)

    regressionStage = "chapter-2-first-real-section"
    VTRegressionAppendSectionHeading testDocument, "Section One", True
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, sectionOneFirstId, latexBase64, 96!, 28!)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, sectionOneSecondId, latexBase64, 96!, 28!)

    regressionStage = "verify-zero-and-real-section-prefixes"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, preChapterFormulaId, 1, "0.0"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, chapterOneFirstId, 1, "1.0"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, chapterOneSecondId, 2, "1.0"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, chapterTwoFirstId, 1, "2.0"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, chapterTwoSecondId, 2, "2.0"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, sectionOneFirstId, 1, "2.1"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, sectionOneSecondId, 2, "2.1"

    regressionStage = "explicit-update-retains-zero-section-prefixes"
    VTUpdateEquationNumbersCore testDocument, True
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, preChapterFormulaId, 1, "0.0"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, chapterOneFirstId, 1, "1.0"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, chapterOneSecondId, 2, "1.0"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, chapterTwoFirstId, 1, "2.0"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, chapterTwoSecondId, 2, "2.0"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, sectionOneFirstId, 1, "2.1"
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, sectionOneSecondId, 2, "2.1"

    resultText = "PASS" & vbLf & _
        "initial=0.0.1,1.0.1,1.0.2,2.0.1,2.0.2,2.1.1,2.1.2" & vbLf & _
        "zeroSectionScope=PASS" & vbLf & _
        "standardHeading1And2=PASS" & vbLf & _
        "customOutlineLevel1=PASS" & vbLf & _
        "explicitUpdateRepair=PASS" & vbLf
    VTWriteTextAtomic resultPath, resultText
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    resultText = "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    VTWriteTextAtomic resultPath, resultText
    ' Preserve the failed document for popup/process/range inspection.
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word zero-section numbering regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordMixedNativeImageChapterAppendRegression()
    Const fixtureFormulaId As String = _
        "11111111-1111-4111-8111-111111111111"
    Const latexBase64 As String = "eF8x"

    Dim testDocument As Document
    Dim formulaRange As Range
    Dim formulaShape As InlineShape
    Dim fixtureRoot As String
    Dim nativeDocumentPath As String
    Dim ommlBase64 As String
    Dim formulaId As String
    Dim firstImageId As String
    Dim secondImageId As String
    Dim resultPath As String
    Dim resultText As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    nativeDocumentPath = _
        VTApplicationSupportRoot() & "/NativeDocuments/" & _
        fixtureFormulaId & ".docx"
    ommlBase64 = VTReadText( _
        fixtureRoot & "/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    resultPath = fixtureRoot & _
        "/word-mixed-native-image-chapter-append-regression-result.txt"
    If Not VTPathFileExists(nativeDocumentPath) Or Len(ommlBase64) = 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The mixed native/image chapter regression fixture is missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_CHAPTER, "-"

    regressionStage = "chapter-1-mixed-native-image"
    VTRegressionAppendChapterHeading testDocument, "Chapter One", True
    formulaId = VTRegressionPerformanceFormulaId(920000001)
    Set formulaRange = VTRegressionCreateNumberedNative( _
        testDocument, formulaId, latexBase64, ommlBase64, _
        nativeDocumentPath)
    VTVerifyNumberedFormulaIntegrity testDocument, formulaId, 1
    If VTEquationNumberTextForFormula(testDocument, formulaId) <> _
       "1" & VTEquationNumberingDisplaySeparator(testDocument) & "1" Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The chapter-1 native OMML number is not 1-1."
    End If
    firstImageId = VTRegressionPerformanceFormulaId(920000101)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, firstImageId, latexBase64, 96!, 28!)
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, firstImageId, 1, "1"

    regressionStage = "chapter-2-mixed-native-image"
    VTRegressionAppendChapterHeading testDocument, "Chapter Two", False
    formulaId = VTRegressionPerformanceFormulaId(920000002)
    Set formulaRange = VTRegressionCreateNumberedNative( _
        testDocument, formulaId, latexBase64, ommlBase64, _
        nativeDocumentPath)
    VTVerifyNumberedFormulaIntegrity testDocument, formulaId, 1
    If VTEquationNumberTextForFormula(testDocument, formulaId) <> _
       "2" & VTEquationNumberingDisplaySeparator(testDocument) & "1" Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The chapter-2 native OMML number is not 2-1."
    End If
    secondImageId = VTRegressionPerformanceFormulaId(920000102)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, secondImageId, latexBase64, 96!, 28!)
    VTRegressionAssertChapterNumberOrdinal _
        testDocument, secondImageId, 1, "2"

    regressionStage = "verify-four-mixed-formulas"
    If VTCountManagedEquationSequences(testDocument) <> 4 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The mixed native/image chapter regression sequence inventory is incomplete."
    End If

    resultText = "PASS" & vbLf & _
        "chapter1=1-1(native),1-1(image)" & vbLf & _
        "chapter2=2-1(native),2-1(image)" & vbLf & _
        "independentSequenceStreams=PASS" & vbLf & _
        "customOutlineLevel1=PASS" & vbLf
    VTWriteTextAtomic resultPath, resultText
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    resultText = "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    VTWriteTextAtomic resultPath, resultText
    ' Preserve the failed document for popup/process/range inspection.
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word mixed native/image chapter append regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Private Function VTRegressionMiddleReferenceResultPath() As String
    VTRegressionMiddleReferenceResultPath = VTApplicationSupportRoot() & _
        "/Tests/word-middle-reference-production-regression-result.txt"
End Function

Private Sub VTRegressionAssertMiddleReferenceState( _
    ByVal documentObject As Document, _
    ByVal expectedBodyReferenceCount As Long)

    Dim formulaIds As Variant
    Dim expectedIds(1 To 5) As String
    Dim candidateField As Field
    Dim helperParagraph As Range
    Dim sequenceBookmarkName As String
    Dim targetBookmarkName As String
    Dim expectedReferenceTarget As String
    Dim expectedReferenceResult As String
    Dim bodyReferenceCount As Long
    Dim sequenceCount As Long
    Dim itemIndex As Long
    Dim assertionStage As String
    Dim assertionErrorNumber As Long
    Dim assertionErrorDescription As String

    On Error GoTo AssertionFailed
    assertionStage = "validate-document"
    If documentObject Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The middle-reference regression document is missing."
    End If
    expectedIds(1) = VTRegressionPerformanceFormulaId(930000001)
    expectedIds(2) = VTRegressionPerformanceFormulaId(930000099)
    expectedIds(3) = VTRegressionPerformanceFormulaId(930000002)
    expectedIds(4) = VTRegressionPerformanceFormulaId(930000003)
    expectedIds(5) = VTRegressionPerformanceFormulaId(930000004)

    assertionStage = "resolve-physical-formula-order"
    formulaIds = VTValidNumberedFormulaIds(documentObject)
    If VTVariantArrayCount(formulaIds) <> 5 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The middle-reference regression lost a numbered formula" & _
            " [live=" & CStr(VTVariantArrayCount(formulaIds)) & "]."
    End If

    assertionStage = "verify-formula-identities-and-ordinals"
    For itemIndex = 1 To 5
        If StrComp(CStr(formulaIds(itemIndex)), expectedIds(itemIndex), _
           vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The middle-reference formula identity order changed" & _
                " [index=" & CStr(itemIndex) & _
                "; actual=" & CStr(formulaIds(itemIndex)) & _
                "; expected=" & expectedIds(itemIndex) & "]."
        End If
        VTVerifyNumberedFormulaIntegrity _
            documentObject, expectedIds(itemIndex), itemIndex
    Next itemIndex

    assertionStage = "verify-first-reference-target"
    expectedReferenceTarget = _
        VTEquationSequenceNumberBookmarkName(expectedIds(1))
    If Not documentObject.Bookmarks.Exists(expectedReferenceTarget) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The middle-reference first Equation lost its VT_N_ target."
    End If
    expectedReferenceResult = Trim$(documentObject.Bookmarks( _
        expectedReferenceTarget).Range.Text)

    assertionStage = "verify-seq-and-body-ref-inventory"
    For Each candidateField In documentObject.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            sequenceCount = sequenceCount + 1
            sequenceBookmarkName = VTSequenceBookmarkNameForField( _
                documentObject, candidateField)
            If Len(sequenceBookmarkName) = 0 Then
                Err.Raise vbObjectError + 7598, "VisualTeX", _
                    "An unbookmarked Equation SEQ helper survived" & _
                    " [range=" & CStr(candidateField.Result.Start) & _
                    "-" & CStr(candidateField.Result.End) & "]."
            End If
            Set helperParagraph = _
                candidateField.Result.Paragraphs(1).Range.Duplicate
            If Not VTHelperParagraphOwnsNativeEquationSequence( _
               helperParagraph) Then
                Err.Raise vbObjectError + 7598, "VisualTeX", _
                    "An Equation SEQ helper was polluted by another field" & _
                    " [bookmark=" & sequenceBookmarkName & _
                    "; fields=" & CStr(helperParagraph.Fields.Count) & "]."
            End If
        ElseIf candidateField.Type = wdFieldRef And _
               Not candidateField.Result.Information(wdWithInTable) And _
               candidateField.Result.Paragraphs(1).Range.InlineShapes.Count = 0 And _
               candidateField.Result.Paragraphs(1).Range.OMaths.Count = 0 Then
            bodyReferenceCount = bodyReferenceCount + 1
            targetBookmarkName = VTReferenceTargetBookmarkName( _
                candidateField.Code.Text)
            If expectedBodyReferenceCount > 0 Then
                If StrComp(targetBookmarkName, expectedReferenceTarget, _
                   vbTextCompare) <> 0 Or _
                   Trim$(candidateField.Result.Text) <> _
                       expectedReferenceResult Then
                    Err.Raise vbObjectError + 7598, "VisualTeX", _
                        "The middle-reference body REF changed identity" & _
                        " [target=" & targetBookmarkName & _
                        "; result=" & candidateField.Result.Text & _
                        "; expectedTarget=" & expectedReferenceTarget & _
                        "; expectedResult=" & expectedReferenceResult & "]."
                End If
            End If
        End If
    Next candidateField

    If sequenceCount <> 5 Or _
       VTCountManagedEquationSequences(documentObject) <> 5 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The middle-reference SEQ inventory is inconsistent" & _
            " [all=" & CStr(sequenceCount) & _
            "; managed=" & _
                CStr(VTCountManagedEquationSequences(documentObject)) & "]."
    End If
    If bodyReferenceCount <> expectedBodyReferenceCount Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The middle-reference body REF count changed" & _
            " [actual=" & CStr(bodyReferenceCount) & _
            "; expected=" & CStr(expectedBodyReferenceCount) & "]."
    End If
    Exit Sub

AssertionFailed:
    assertionErrorNumber = Err.Number
    assertionErrorDescription = Err.Description
    If assertionErrorNumber = 0 Then
        assertionErrorNumber = vbObjectError + 7598
    End If
    If Len(assertionErrorDescription) = 0 Then
        assertionErrorDescription = _
            "The regression assertion failed without a Word error description."
    End If
    Err.Raise assertionErrorNumber, _
        "VisualTeX Word middle-reference state assertion", _
        assertionStage & ": " & assertionErrorDescription
End Sub

Private Function VTWordMiddleReferenceRegressionDocument() As Document
    Dim candidateDocument As Document
    Dim matchedDocument As Document
    Dim regressionMarker As String

    For Each candidateDocument In Documents
        regressionMarker = ""
        If VTTryGetDocumentVariable( _
           candidateDocument, "VT_MiddleReferenceProductionRegression", _
           regressionMarker) Then
            If regressionMarker = "v1" Then
                If Not matchedDocument Is Nothing Then
                    Err.Raise vbObjectError + 7598, "VisualTeX", _
                        "More than one middle-reference regression document is open."
                End If
                Set matchedDocument = candidateDocument
            End If
        End If
    Next candidateDocument
    If matchedDocument Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The middle-reference regression document is not open."
    End If
    Set VTWordMiddleReferenceRegressionDocument = matchedDocument
End Function

Private Function VTWordMiddleReferenceRegressionTargetRange( _
    ByVal documentObject As Document) As Range

    Const targetBookmarkName As String = "VT_MR_Target"

    If documentObject Is Nothing Or _
       Not documentObject.Bookmarks.Exists(targetBookmarkName) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The middle-reference regression insertion target is missing."
    End If
    Set VTWordMiddleReferenceRegressionTargetRange = _
        documentObject.Bookmarks(targetBookmarkName).Range.Duplicate
End Function

Public Sub VisualTeX_DumpWordMiddleReferenceProductionRegressionOpenXml()
    Dim documentObject As Document
    Dim outputPath As String

    Set documentObject = VTWordMiddleReferenceRegressionDocument()
    documentObject.Activate
    outputPath = VTApplicationSupportRoot() & _
        "/Tests/word-middle-reference-production-openxml.xml"
    VTWriteTextAtomic outputPath, documentObject.Content.WordOpenXML
End Sub

Public Sub VisualTeX_PrepareWordMiddleReferenceProductionRegression()
    Const fixtureFormulaId As String = _
        "11111111-1111-4111-8111-111111111111"
    Const latexBase64 As String = "eF8x"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim insertionRange As Range
    Dim fixtureRoot As String
    Dim nativeDocumentPath As String
    Dim ommlBase64 As String
    Dim targetStart As Long
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    nativeDocumentPath = VTApplicationSupportRoot() & _
        "/NativeDocuments/" & fixtureFormulaId & ".docx"
    ommlBase64 = VTReadText( _
        fixtureRoot & "/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    If Not VTPathFileExists(nativeDocumentPath) Or Len(ommlBase64) = 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The middle-reference native regression fixture is missing."
    End If

    regressionStage = "create-document"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_CHAPTER, "-"

    regressionStage = "create-leading-image"
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, VTRegressionPerformanceFormulaId(930000001), _
        latexBase64, 96!, 28!)

    regressionStage = "create-three-native-formulas"
    Set formulaRange = VTRegressionCreateNumberedNative( _
        testDocument, VTRegressionPerformanceFormulaId(930000002), _
        latexBase64, ommlBase64, nativeDocumentPath)
    Set formulaRange = VTRegressionCreateNumberedNative( _
        testDocument, VTRegressionPerformanceFormulaId(930000003), _
        latexBase64, ommlBase64, nativeDocumentPath)
    Set formulaRange = VTRegressionCreateNumberedNative( _
        testDocument, VTRegressionPerformanceFormulaId(930000004), _
        latexBase64, ommlBase64, nativeDocumentPath)

    regressionStage = "insert-middle-image"
    Set insertionRange = VTRegressionInsertBlankParagraphBeforeFormula( _
        testDocument, VTRegressionPerformanceFormulaId(930000002))
    Set formulaShape = VTRegressionCreateNumberedImageAtCaret( _
        testDocument, insertionRange, _
        VTRegressionPerformanceFormulaId(930000099), _
        latexBase64, 96!, 28!)

    regressionStage = "verify-middle-insertion"
    VTRegressionAssertMiddleReferenceState testDocument, 0

    regressionStage = "select-reference-target"
    VTSetDocumentVariable _
        testDocument, "VT_MiddleReferenceProductionRegression", "v1"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    ' VTAppendRegressionParagraph can inherit the compact hidden-helper paragraph
    ' formatting when the final numbered formula is native OMML. The regression
    ' target is ordinary body text, so normalize that fresh paragraph explicitly
    ' before measuring the insertion Bookmark. This keeps the fixture faithful to
    ' the real Ribbon contract instead of manufacturing a false damaged helper.
    VTNormalizePlainWordParagraph insertionRange.Paragraphs(1).Range
    targetStart = insertionRange.Start
    insertionRange.InsertAfter "VT_MIDDLE_REFERENCE_TARGET "
    Set insertionRange = testDocument.Range( _
        Start:=targetStart + Len("VT_MIDDLE_REFERENCE_TARGET "), _
        End:=targetStart + Len("VT_MIDDLE_REFERENCE_TARGET "))
    If testDocument.Bookmarks.Exists("VT_MR_Target") Then
        testDocument.Bookmarks("VT_MR_Target").Delete
    End If
    testDocument.Bookmarks.Add _
        name:="VT_MR_Target", Range:=insertionRange.Duplicate
    testDocument.Activate
    insertionRange.Select

    VTWriteTextAtomic VTRegressionMiddleReferenceResultPath(), _
        "PREPARED" & vbLf & _
        "formulaOrder=1,2,3,4,5" & vbLf & _
        "bodyReferences=0" & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic VTRegressionMiddleReferenceResultPath(), _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            regressionErrorDescription, vbCr, " "), vbLf, " ") & vbLf
    If Not testDocument Is Nothing Then
        testDocument.Saved = True
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word middle-reference production regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_VerifyWordMiddleReferencePickerCancelRegression()
    Dim documentObject As Document
    Dim targetRange As Range
    Dim targetParagraph As Range

    On Error GoTo RegressionFailed
    Set documentObject = VTWordMiddleReferenceRegressionDocument()
    documentObject.Activate
    VTRegressionAssertMiddleReferenceState documentObject, 0
    Set targetRange = VTWordMiddleReferenceRegressionTargetRange(documentObject)
    Set targetParagraph = targetRange.Paragraphs(1).Range.Duplicate
    If VTParagraphHasNumberHelperGeometry(targetParagraph) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The Equation picker target Bookmark moved into an internal helper."
    End If
    VTWriteTextAtomic VTRegressionMiddleReferenceResultPath(), _
        "PICKER_CANCEL_PASS" & vbLf & _
        "formulaOrder=1,2,3,4,5" & vbLf & _
        "bodyReferences=0" & vbLf
    Exit Sub
RegressionFailed:
    VTWriteTextAtomic VTRegressionMiddleReferenceResultPath(), _
        "FAIL" & vbLf & _
        "stage=picker-cancel-verify" & vbLf & _
        "errorNumber=" & CStr(Err.Number) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            Err.Description, vbCr, " "), vbLf, " ") & vbLf
    Err.Raise Err.Number, _
        "VisualTeX Word middle-reference picker cancel regression", _
        Err.Description
End Sub

Private Sub VTSetWordMiddleReferenceRegressionSelection( _
    ByVal selectionValue As String)

    Dim documentObject As Document
    Dim targetRange As Range

    Set documentObject = VTWordMiddleReferenceRegressionDocument()
    documentObject.Activate
    Set targetRange = VTWordMiddleReferenceRegressionTargetRange(documentObject)
    targetRange.Select
    VTSetDocumentVariable _
        documentObject, "VT_MiddleReferenceRegressionSelection", _
        selectionValue
End Sub

Public Sub VisualTeX_SetWordMiddleReferenceRegressionSelection()
    VTSetWordMiddleReferenceRegressionSelection "1"
End Sub

Public Sub VisualTeX_SetWordMiddleReferenceRegressionCancelSelection()
    ' A single space is intentionally non-empty in the Document Variable so
    ' VTTryGetDocumentVariable resolves the deterministic test hook. The public
    ' production command then trims it to empty and follows the same cancel path
    ' as a user dismissing InputBox, without depending on macOS Accessibility.
    VTSetWordMiddleReferenceRegressionSelection " "
End Sub

Public Sub VisualTeX_VerifyWordMiddleReferenceInsertedRegression()
    Dim documentObject As Document

    On Error GoTo RegressionFailed
    Set documentObject = VTWordMiddleReferenceRegressionDocument()
    documentObject.Activate
    VTRegressionAssertMiddleReferenceState documentObject, 1
    VTWriteTextAtomic VTRegressionMiddleReferenceResultPath(), _
        "REFERENCE_INSERT_PASS" & vbLf & _
        "formulaOrder=1,2,3,4,5" & vbLf & _
        "bodyReferences=1" & vbLf
    Exit Sub
RegressionFailed:
    VTWriteTextAtomic VTRegressionMiddleReferenceResultPath(), _
        "FAIL" & vbLf & _
        "stage=reference-insert-verify" & vbLf & _
        "errorNumber=" & CStr(Err.Number) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            Err.Description, vbCr, " "), vbLf, " ") & vbLf
    Err.Raise Err.Number, _
        "VisualTeX Word middle-reference insertion regression", _
        Err.Description
End Sub

Public Sub VisualTeX_RunWordMiddleReferenceInsertProductionRegression()
    Dim documentObject As Document
    Dim insertionRange As Range
    Dim insertedRange As Range
    Dim crossReferenceItems As Variant
    Dim expectedReferenceText As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed

    regressionStage = "prepare-middle-insertion"
    VisualTeX_PrepareWordMiddleReferenceProductionRegression
    Set documentObject = VTWordMiddleReferenceRegressionDocument()
    Set insertionRange = VTWordMiddleReferenceRegressionTargetRange( _
        documentObject)

    regressionStage = "production-build-reference-list"
    crossReferenceItems = VTEquationNumberCrossReferenceItems(documentObject)
    If VTVariantArrayCount(crossReferenceItems) <> 5 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The production Equation picker list does not contain five live formulas."
    End If
    VTRegressionAssertMiddleReferenceState documentObject, 0

    regressionStage = "production-insert-reference"
    expectedReferenceText = "(" & Trim$(documentObject.Bookmarks( _
        VTEquationSequenceNumberBookmarkName( _
            VTRegressionPerformanceFormulaId(930000001))).Range.Text) & ")"
    Set insertedRange = VTInsertEquationNumberReferenceAtRange( _
        insertionRange, 1, True)
    If insertedRange Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The production Equation reference insertion returned no Range."
    End If
    If insertedRange.Text <> expectedReferenceText Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The production Equation reference insertion changed the live " & _
            "VT_N_ number [actual=" & insertedRange.Text & _
            "; expected=" & expectedReferenceText & "]."
    End If

    regressionStage = "capture-production-reference"
    VTWriteTextAtomic _
        VTApplicationSupportRoot() & _
            "/Tests/word-middle-reference-after-public-insert.xml", _
        documentObject.Content.WordOpenXML

    regressionStage = "verify-production-reference"
    VTRegressionAssertMiddleReferenceState documentObject, 1

    VTWriteTextAtomic VTRegressionMiddleReferenceResultPath(), _
        "PASS" & vbLf & _
        "stage=production-reference" & vbLf & _
        "formulaOrder=1,2,3,4,5" & vbLf & _
        "bodyReferences=1" & vbLf & _
        "sequenceInventory=5" & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    If regressionErrorNumber = 0 Then
        regressionErrorNumber = vbObjectError + 7598
    End If
    If Len(regressionErrorDescription) = 0 Then
        regressionErrorDescription = _
            "The production reference regression failed without a Word error description."
    End If
    On Error Resume Next
    VTWriteTextAtomic VTRegressionMiddleReferenceResultPath(), _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            regressionErrorDescription, vbCr, " "), vbLf, " ") & vbLf
    If Not documentObject Is Nothing Then
        documentObject.Saved = True
        documentObject.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word middle-reference production-path regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_SelectWordMiddleReferenceInsertedFormula()
    Dim documentObject As Document
    Dim formulaRange As Range

    On Error GoTo RegressionFailed
    Set documentObject = VTWordMiddleReferenceRegressionDocument()
    documentObject.Activate
    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, VTRegressionPerformanceFormulaId(930000099))
    If formulaRange Is Nothing Or formulaRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The inserted middle image formula is missing before edit."
    End If
    formulaRange.InlineShapes(1).Range.Select
    Exit Sub
RegressionFailed:
    Err.Raise Err.Number, _
        "VisualTeX Word middle-reference edit selection regression", _
        Err.Description
End Sub

Public Sub VisualTeX_VerifyWordMiddleReferenceAfterEditRegression()
    Dim documentObject As Document

    On Error GoTo RegressionFailed
    Set documentObject = VTWordMiddleReferenceRegressionDocument()
    documentObject.Activate
    VTRegressionAssertMiddleReferenceState documentObject, 1
    VTWriteTextAtomic VTRegressionMiddleReferenceResultPath(), _
        "EDIT_PASS" & vbLf & _
        "formulaOrder=1,2,3,4,5" & vbLf & _
        "bodyReferences=1" & vbLf
    Exit Sub
RegressionFailed:
    VTWriteTextAtomic VTRegressionMiddleReferenceResultPath(), _
        "FAIL" & vbLf & _
        "stage=edit-verify" & vbLf & _
        "errorNumber=" & CStr(Err.Number) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            Err.Description, vbCr, " "), vbLf, " ") & vbLf
    Err.Raise Err.Number, _
        "VisualTeX Word middle-reference edit regression", _
        Err.Description
End Sub

Public Sub VisualTeX_VerifyWordMiddleReferenceAfterUpdateRegression()
    Dim documentObject As Document

    On Error GoTo RegressionFailed
    Set documentObject = VTWordMiddleReferenceRegressionDocument()
    documentObject.Activate
    VTRegressionAssertMiddleReferenceState documentObject, 1
    VTWriteTextAtomic VTRegressionMiddleReferenceResultPath(), _
        "PASS" & vbLf & _
        "pickerCancel=PASS" & vbLf & _
        "referenceInsert=PASS" & vbLf & _
        "edit=PASS" & vbLf & _
        "explicitUpdate=PASS" & vbLf & _
        "formulaOrder=1,2,3,4,5" & vbLf & _
        "bodyReferences=1" & vbLf & _
        "sequenceInventory=5" & vbLf
    Exit Sub
RegressionFailed:
    VTWriteTextAtomic VTRegressionMiddleReferenceResultPath(), _
        "FAIL" & vbLf & _
        "stage=explicit-update-verify" & vbLf & _
        "errorNumber=" & CStr(Err.Number) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            Err.Description, vbCr, " "), vbLf, " ") & vbLf
    Err.Raise Err.Number, _
        "VisualTeX Word middle-reference update regression", _
        Err.Description
End Sub

Private Function VTRegressionPerformanceFormulaId( _
    ByVal itemIndex As Long) As String

    If itemIndex < 1 Or itemIndex > 999999999 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The numbering performance formula index is invalid."
    End If
    VTRegressionPerformanceFormulaId = _
        "70000000-0000-4000-8000-" & _
        Right$("000000000000" & CStr(itemIndex), 12)
End Function

Public Sub VisualTeX_RunWordSequentialNumberingPerformanceRegression()
    Const formulaCount As Long = 60
    Const latexBase64 As String = "eF4yK3leMg"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim formulaParagraph As Range
    Dim legacyTable As Table
    Dim insertSeconds(1 To formulaCount) As Double
    Dim resultPath As String
    Dim resultText As String
    Dim formulaId As String
    Dim startedAt As Single
    Dim afterThirtyAverage As Double
    Dim tailAverage As Double
    Dim editAfterThirtySeconds As Double
    Dim editAfterSixtySeconds As Double
    Dim migrationAfterThirtySeconds As Double
    Dim migrationAfterSixtySeconds As Double
    Dim numberCreated As Boolean
    Dim itemIndex As Long
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-sequential-numbering-performance-regression-result.txt"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_MODE_VARIABLE, _
        VT_WORD_NUMBERING_MODE_SEQUENCE
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_SEPARATOR_VARIABLE, "."

    For itemIndex = 1 To 30
        formulaId = VTRegressionPerformanceFormulaId(itemIndex)
        startedAt = Timer
        Set formulaShape = VTRegressionCreateNumberedImage( _
            testDocument, formulaId, latexBase64, 96!, 28!)
        insertSeconds(itemIndex) = _
            VTTimerElapsedSeconds(startedAt, Timer)
    Next itemIndex
    editAfterThirtySeconds = VTRegressionNumberingRefreshSeconds( _
        testDocument, VTRegressionPerformanceFormulaId(30), 3)
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, VTRegressionPerformanceFormulaId(30))
    Set formulaParagraph = formulaRange.Paragraphs(1).Range.Duplicate
    Set legacyTable = formulaParagraph.ConvertToTable( _
        Separator:=wdSeparateByTabs, NumRows:=1, NumColumns:=3)
    Set formulaShape = legacyTable.Cell(1, 2).Range.InlineShapes(1)
    startedAt = Timer
    numberCreated = False
    Set formulaRange = VTEnsureImageEquationNumber( _
        formulaShape, formulaShape.Height, _
        VTRegressionPerformanceFormulaId(30), _
        "performance legacy-table migration", numberCreated)
    migrationAfterThirtySeconds = _
        VTTimerElapsedSeconds(startedAt, Timer)
    If formulaShape.Range.Information(wdWithInTable) Or _
       testDocument.Tables.Count <> 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The 30-formula performance migration retained a numbered image table."
    End If

    For itemIndex = 31 To formulaCount
        formulaId = VTRegressionPerformanceFormulaId(itemIndex)
        startedAt = Timer
        Set formulaShape = VTRegressionCreateNumberedImage( _
            testDocument, formulaId, latexBase64, 96!, 28!)
        insertSeconds(itemIndex) = _
            VTTimerElapsedSeconds(startedAt, Timer)
    Next itemIndex
    editAfterSixtySeconds = VTRegressionNumberingRefreshSeconds( _
        testDocument, VTRegressionPerformanceFormulaId(60), 3)
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, VTRegressionPerformanceFormulaId(60))
    Set formulaParagraph = formulaRange.Paragraphs(1).Range.Duplicate
    Set legacyTable = formulaParagraph.ConvertToTable( _
        Separator:=wdSeparateByTabs, NumRows:=1, NumColumns:=3)
    Set formulaShape = legacyTable.Cell(1, 2).Range.InlineShapes(1)
    startedAt = Timer
    numberCreated = False
    Set formulaRange = VTEnsureImageEquationNumber( _
        formulaShape, formulaShape.Height, _
        VTRegressionPerformanceFormulaId(60), _
        "performance legacy-table migration", numberCreated)
    migrationAfterSixtySeconds = _
        VTTimerElapsedSeconds(startedAt, Timer)
    If formulaShape.Range.Information(wdWithInTable) Or _
       testDocument.Tables.Count <> 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The 60-formula performance migration retained a numbered image table."
    End If

    For itemIndex = 31 To 35
        afterThirtyAverage = afterThirtyAverage + insertSeconds(itemIndex)
    Next itemIndex
    afterThirtyAverage = afterThirtyAverage / 5#
    For itemIndex = formulaCount - 4 To formulaCount
        tailAverage = tailAverage + insertSeconds(itemIndex)
    Next itemIndex
    tailAverage = tailAverage / 5#

    If tailAverage > afterThirtyAverage * 1.5 + 0.25 Or _
       editAfterSixtySeconds > editAfterThirtySeconds * 1.5 + 0.25 Or _
       migrationAfterSixtySeconds > _
           migrationAfterThirtySeconds * 1.5 + 0.25 Then
        resultText = "FAIL" & vbLf
    Else
        resultText = "PASS" & vbLf
    End If
    resultText = resultText & _
        "formulaCount=" & CStr(formulaCount) & vbLf & _
        "afterThirtyAvgSeconds=" & _
            Format$(afterThirtyAverage, "0.000") & vbLf & _
        "tailAvgSeconds=" & Format$(tailAverage, "0.000") & vbLf & _
        "growthRatio=" & _
            Format$(tailAverage / afterThirtyAverage, "0.000") & vbLf & _
        "item30Seconds=" & Format$(insertSeconds(30), "0.000") & vbLf & _
        "item60Seconds=" & Format$(insertSeconds(60), "0.000") & vbLf
    resultText = resultText & _
        "editAfterThirtySeconds=" & _
            Format$(editAfterThirtySeconds, "0.000") & vbLf & _
        "editAfterSixtySeconds=" & _
            Format$(editAfterSixtySeconds, "0.000") & vbLf & _
        "editGrowthRatio=" & _
            Format$(editAfterSixtySeconds / editAfterThirtySeconds, _
                "0.000") & vbLf & _
        "migrationAfterThirtySeconds=" & _
            Format$(migrationAfterThirtySeconds, "0.000") & vbLf & _
        "migrationAfterSixtySeconds=" & _
            Format$(migrationAfterSixtySeconds, "0.000") & vbLf & _
        "migrationGrowthRatio=" & _
            Format$(migrationAfterSixtySeconds / _
                migrationAfterThirtySeconds, "0.000") & vbLf & _
        "tablesAfterMigration=" & CStr(testDocument.Tables.Count) & vbLf
    VTWriteTextAtomic resultPath, resultText
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    resultText = "FAIL" & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    VTWriteTextAtomic resultPath, resultText
    ' Preserve the failed document for popup/process/range inspection.
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word sequential numbering performance regression", _
        regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordNativeInsertFileProbe()
    Const fixtureFormulaId As String = _
        "11111111-1111-4111-8111-111111111111"
    Const latexBase64 As String = "eF8x"

    Dim testDocument As Document
    Dim insertionRange As Range
    Dim probeRange As Range
    Dim equationRange As Range
    Dim numberRange As Range
    Dim typedRange As Range
    Dim nativeMath As OMath
    Dim fixtureRoot As String
    Dim nativeDocumentPath As String
    Dim ommlBase64 As String
    Dim formulaId As String
    Dim resultPath As String
    Dim resultText As String
    Dim startedAt As Single
    Dim elapsedSeconds As Double
    Dim numberingSeconds As Double
    Dim insertionStart As Long
    Dim textStart As Long
    Dim numberCreated As Boolean
    Dim probeErrorNumber As Long
    Dim probeErrorDescription As String

    On Error GoTo ProbeFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    nativeDocumentPath = _
        VTApplicationSupportRoot() & "/NativeDocuments/" & _
        fixtureFormulaId & ".docx"
    ommlBase64 = VTReadText( _
        fixtureRoot & "/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    formulaId = VTRegressionPerformanceFormulaId(1)
    resultPath = fixtureRoot & _
        "/word-native-insert-file-probe-result.txt"
    If Not VTPathFileExists(nativeDocumentPath) Or Len(ommlBase64) = 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The native InsertFile probe fixture is missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    Set insertionRange = testDocument.Content.Duplicate
    insertionRange.Collapse wdCollapseStart
    insertionStart = insertionRange.Start
    startedAt = Timer
    insertionRange.InsertFile _
        FileName:=nativeDocumentPath, _
        ConfirmConversions:=False, Link:=False, Attachment:=False
    elapsedSeconds = VTTimerElapsedSeconds(startedAt, Timer)
    Set probeRange = testDocument.Range( _
        Start:=insertionStart, End:=testDocument.Content.End)

    If probeRange.OMaths.Count = 1 Then
        Set nativeMath = VTNativeMathRoot(probeRange.OMaths(1))
    End If
    If nativeMath Is Nothing Or Not VTOMathHasMeaningfulContent(nativeMath) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "InsertFile did not produce one meaningful native Equation."
    End If

    Set equationRange = VTFinalizeInlineNativeEquation( _
        nativeMath.Range.Duplicate)
    VTSetWordLatexPayload testDocument, formulaId, latexBase64
    VTSetWordOmmlPayload testDocument, formulaId, ommlBase64
    VTSetWordMetadataPayload testDocument, formulaId, VT_METADATA_PREFIX & "e30"
    VTSetWordFormulaFormat testDocument, formulaId, "block", True
    startedAt = Timer
    numberCreated = False
    Set numberRange = VTEnsureNativeEquationNumber( _
        equationRange, 48#, formulaId, _
        "InsertFile probe native formula", numberCreated)
    numberingSeconds = VTTimerElapsedSeconds(startedAt, Timer)
    If Not numberCreated Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "InsertFile probe did not create a fresh Equation number."
    End If
    VTVerifyNumberedFormulaIntegrity testDocument, formulaId, 1

    Set equationRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    If equationRange Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "InsertFile probe lost the numbered native Equation."
    End If
    VTPlaceCaretAfterDisplayFormula equationRange, formulaId
    textStart = Selection.Start
    Selection.TypeText Text:="InsertFile continuation"
    Set typedRange = testDocument.Range( _
        Start:=textStart, End:=Selection.Start)
    If typedRange.Text <> "InsertFile continuation" Or _
       typedRange.Information(wdWithInTable) Or _
       typedRange.Font.Hidden <> False Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "InsertFile probe did not return to ordinary body text."
    End If

    resultText = "PASS" & vbLf & _
        "elapsedSeconds=" & Format$(elapsedSeconds, "0.000") & vbLf & _
        "numberingSeconds=" & Format$(numberingSeconds, "0.000") & vbLf
    resultText = resultText & _
        "mathCount=" & CStr(probeRange.OMaths.Count) & vbLf & _
        "paragraphCount=" & CStr(testDocument.Paragraphs.Count) & vbLf & _
        "tableCount=" & CStr(testDocument.Tables.Count) & vbLf
    resultText = resultText & _
        "inlineShapeCount=" & _
            CStr(testDocument.InlineShapes.Count) & vbLf & _
        "sectionCount=" & CStr(testDocument.Sections.Count) & vbLf & _
        "continuationInTable=" & _
            CStr(typedRange.Information(wdWithInTable)) & vbLf
    If Not nativeMath Is Nothing Then
        resultText = resultText & _
            "mathMeaningful=" & _
                CStr(VTOMathHasMeaningfulContent(nativeMath)) & vbLf & _
            "mathStart=" & CStr(nativeMath.Range.Start) & vbLf & _
            "mathEnd=" & CStr(nativeMath.Range.End) & vbLf
    End If
    VTWriteTextAtomic resultPath, resultText
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

ProbeFailed:
    probeErrorNumber = Err.Number
    probeErrorDescription = Err.Description
    On Error Resume Next
    resultText = "FAIL" & vbLf & _
        "errorNumber=" & CStr(probeErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(probeErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    VTWriteTextAtomic resultPath, resultText
    If Not testDocument Is Nothing Then
        testDocument.Saved = True
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise probeErrorNumber, _
        "VisualTeX Word native InsertFile probe", probeErrorDescription
End Sub

Public Sub VisualTeX_RunWordSequentialNativeNumberingPerformanceRegression()
    Const formulaCount As Long = 45
    Const fixtureFormulaId As String = _
        "11111111-1111-4111-8111-111111111111"
    Const latexBase64 As String = "eF8x"

    Dim testDocument As Document
    Dim formulaRange As Range
    Dim insertSeconds(1 To formulaCount) As Double
    Dim insertPhaseTimings(1 To formulaCount) As Variant
    Dim fixtureRoot As String
    Dim nativeDocumentPath As String
    Dim ommlBase64 As String
    Dim resultPath As String
    Dim resultText As String
    Dim formulaId As String
    Dim startedAt As Single
    Dim afterThirtyAverage As Double
    Dim tailAverage As Double
    Dim editAfterThirtySeconds As Double
    Dim editAfterFortyFiveSeconds As Double
    Dim earlyPhases(0 To 16) As Double
    Dim latePhases(0 To 16) As Double
    Dim itemIndex As Long
    Dim phaseIndex As Long
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    nativeDocumentPath = _
        VTApplicationSupportRoot() & "/NativeDocuments/" & _
        fixtureFormulaId & ".docx"
    ommlBase64 = VTReadText( _
        fixtureRoot & "/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    resultPath = fixtureRoot & _
        "/word-sequential-native-numbering-performance-regression-result.txt"
    If Not VTPathFileExists(nativeDocumentPath) Or Len(ommlBase64) = 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The native numbering performance fixture is missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."

    For itemIndex = 1 To 30
        formulaId = VTRegressionPerformanceFormulaId(itemIndex)
        startedAt = Timer
        Set formulaRange = VTRegressionCreateNumberedNative( _
            testDocument, formulaId, latexBase64, ommlBase64, _
            nativeDocumentPath, insertPhaseTimings(itemIndex))
        insertSeconds(itemIndex) = _
            VTTimerElapsedSeconds(startedAt, Timer)
    Next itemIndex
    editAfterThirtySeconds = VTRegressionNativeNumberingRefreshSeconds( _
        testDocument, VTRegressionPerformanceFormulaId(30), 3)

    For itemIndex = 31 To formulaCount
        formulaId = VTRegressionPerformanceFormulaId(itemIndex)
        startedAt = Timer
        Set formulaRange = VTRegressionCreateNumberedNative( _
            testDocument, formulaId, latexBase64, ommlBase64, _
            nativeDocumentPath, insertPhaseTimings(itemIndex))
        insertSeconds(itemIndex) = _
            VTTimerElapsedSeconds(startedAt, Timer)
    Next itemIndex
    editAfterFortyFiveSeconds = _
        VTRegressionNativeNumberingRefreshSeconds( _
            testDocument, VTRegressionPerformanceFormulaId(45), 3)

    For itemIndex = 31 To 35
        afterThirtyAverage = afterThirtyAverage + insertSeconds(itemIndex)
        For phaseIndex = 0 To 16
            earlyPhases(phaseIndex) = earlyPhases(phaseIndex) + _
                CDbl(insertPhaseTimings(itemIndex)(phaseIndex))
        Next phaseIndex
    Next itemIndex
    afterThirtyAverage = afterThirtyAverage / 5#
    For itemIndex = formulaCount - 4 To formulaCount
        tailAverage = tailAverage + insertSeconds(itemIndex)
        For phaseIndex = 0 To 16
            latePhases(phaseIndex) = latePhases(phaseIndex) + _
                CDbl(insertPhaseTimings(itemIndex)(phaseIndex))
        Next phaseIndex
    Next itemIndex
    tailAverage = tailAverage / 5#
    For phaseIndex = 0 To 16
        earlyPhases(phaseIndex) = earlyPhases(phaseIndex) / 5#
        latePhases(phaseIndex) = latePhases(phaseIndex) / 5#
    Next phaseIndex

    If tailAverage > afterThirtyAverage * 1.5 + 0.25 Or _
       editAfterFortyFiveSeconds > editAfterThirtySeconds * 1.5 + 0.25 Then
        resultText = "FAIL" & vbLf
    Else
        resultText = "PASS" & vbLf
    End If
    resultText = resultText & _
        "formulaCount=" & CStr(formulaCount) & vbLf & _
        "afterThirtyAvgSeconds=" & _
            Format$(afterThirtyAverage, "0.000") & vbLf & _
        "tailAvgSeconds=" & Format$(tailAverage, "0.000") & vbLf & _
        "growthRatio=" & _
            Format$(tailAverage / afterThirtyAverage, "0.000") & vbLf
    resultText = resultText & _
        "editAfterThirtySeconds=" & _
            Format$(editAfterThirtySeconds, "0.000") & vbLf & _
        "editAfterFortyFiveSeconds=" & _
            Format$(editAfterFortyFiveSeconds, "0.000") & vbLf & _
        "editGrowthRatio=" & _
            Format$(editAfterFortyFiveSeconds / editAfterThirtySeconds, _
                "0.000") & vbLf
    resultText = resultText & _
        "nativeAppendEarlyLate=" & _
            Format$(earlyPhases(0), "0.000") & "/" & _
            Format$(latePhases(0), "0.000") & vbLf & _
        "nativeInsertEarlyLate=" & _
            Format$(earlyPhases(1), "0.000") & "/" & _
            Format$(latePhases(1), "0.000") & vbLf
    resultText = resultText & _
        "nativeMetadataEarlyLate=" & _
            Format$(earlyPhases(2), "0.000") & "/" & _
            Format$(latePhases(2), "0.000") & vbLf & _
        "nativeNumberingEarlyLate=" & _
            Format$(earlyPhases(3), "0.000") & "/" & _
            Format$(latePhases(3), "0.000") & vbLf
    resultText = resultText & _
        "nativeResolveEarlyLate=" & _
            Format$(earlyPhases(4), "0.000") & "/" & _
            Format$(latePhases(4), "0.000") & vbLf
    resultText = resultText & _
        "nativeNumberPrepareEarlyLate=" & _
            Format$(earlyPhases(5), "0.000") & "/" & _
            Format$(latePhases(5), "0.000") & vbLf
    resultText = resultText & _
        "nativeNumberHelperEarlyLate=" & _
            Format$(earlyPhases(6), "0.000") & "/" & _
            Format$(latePhases(6), "0.000") & vbLf
    resultText = resultText & _
        "nativeNumberMarkerEarlyLate=" & _
            Format$(earlyPhases(7), "0.000") & "/" & _
            Format$(latePhases(7), "0.000") & vbLf
    resultText = resultText & _
        "nativeNumberReferenceEarlyLate=" & _
            Format$(earlyPhases(8), "0.000") & "/" & _
            Format$(latePhases(8), "0.000") & vbLf
    resultText = resultText & _
        "nativeNumberBuildEarlyLate=" & _
            Format$(earlyPhases(9), "0.000") & "/" & _
            Format$(latePhases(9), "0.000") & vbLf
    resultText = resultText & _
        "nativeNumberFinalizeEarlyLate=" & _
            Format$(earlyPhases(10), "0.000") & "/" & _
            Format$(latePhases(10), "0.000") & vbLf
    resultText = resultText & _
        "nativeFastPathGateEarlyLate=" & _
            Format$(earlyPhases(11), "0.0") & "/" & _
            Format$(latePhases(11), "0.0") & vbLf
    resultText = resultText & _
        "nativeInsertSetupEarlyLate=" & _
            Format$(earlyPhases(12), "0.000") & "/" & _
            Format$(latePhases(12), "0.000") & vbLf
    resultText = resultText & _
        "nativeInsertStagingEarlyLate=" & _
            Format$(earlyPhases(13), "0.000") & "/" & _
            Format$(latePhases(13), "0.000") & vbLf
    resultText = resultText & _
        "nativeInsertTransferEarlyLate=" & _
            Format$(earlyPhases(14), "0.000") & "/" & _
            Format$(latePhases(14), "0.000") & vbLf
    resultText = resultText & _
        "nativeInsertResolveEarlyLate=" & _
            Format$(earlyPhases(15), "0.000") & "/" & _
            Format$(latePhases(15), "0.000") & vbLf
    resultText = resultText & _
        "nativeInsertFinalizeEarlyLate=" & _
            Format$(earlyPhases(16), "0.000") & "/" & _
            Format$(latePhases(16), "0.000") & vbLf
    VTWriteTextAtomic resultPath, resultText
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    resultText = "FAIL" & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    VTWriteTextAtomic resultPath, resultText
    If Not testDocument Is Nothing Then
        testDocument.Saved = True
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word native sequential numbering performance regression", _
        regressionErrorDescription
End Sub

Private Function VTRegressionNativeNumberingRefreshSeconds( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal repeatCount As Long) As Double

    Dim formulaRange As Range
    Dim numberRange As Range
    Dim startedAt As Single
    Dim numberCreated As Boolean
    Dim itemIndex As Long

    If documentObject Is Nothing Or repeatCount < 1 Then Exit Function
    For itemIndex = 1 To repeatCount
        Set formulaRange = VTNumberedFormulaRangeForId( _
            documentObject, formulaId)
        If formulaRange Is Nothing Or formulaRange.OMaths.Count <> 1 Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The native numbering performance edit target is missing."
        End If
        startedAt = Timer
        numberCreated = False
        Set numberRange = VTEnsureNativeEquationNumber( _
            formulaRange, 48#, formulaId, _
            "native performance edit", numberCreated)
        If numberCreated Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "Editing a native numbered formula recreated its number identity."
        End If
        VTRegressionNativeNumberingRefreshSeconds = _
            VTRegressionNativeNumberingRefreshSeconds + _
            VTTimerElapsedSeconds(startedAt, Timer)
    Next itemIndex
    VTRegressionNativeNumberingRefreshSeconds = _
        VTRegressionNativeNumberingRefreshSeconds / CDbl(repeatCount)
End Function

Private Function VTRegressionNumberingRefreshSeconds( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal repeatCount As Long) As Double

    Dim formulaRange As Range
    Dim formulaShape As InlineShape
    Dim startedAt As Single
    Dim itemIndex As Long

    If documentObject Is Nothing Or repeatCount < 1 Then Exit Function
    For itemIndex = 1 To repeatCount
        Set formulaRange = VTNumberedFormulaRangeForId( _
            documentObject, formulaId)
        If formulaRange Is Nothing Or _
           formulaRange.InlineShapes.Count <> 1 Then
            Err.Raise vbObjectError + 7598, "VisualTeX", _
                "The numbering performance edit target is missing."
        End If
        Set formulaShape = formulaRange.InlineShapes(1)
        startedAt = Timer
        Set formulaRange = VTInsertEquationNumber( _
            formulaShape, formulaId, "performance edit")
        VTRegressionNumberingRefreshSeconds = _
            VTRegressionNumberingRefreshSeconds + _
            VTTimerElapsedSeconds(startedAt, Timer)
    Next itemIndex
    VTRegressionNumberingRefreshSeconds = _
        VTRegressionNumberingRefreshSeconds / CDbl(repeatCount)
End Function

Public Sub VisualTeX_RunWordNumberingPerformanceRegression()
    Const formulaCount As Long = 36
    Const latexBase64 As String = "eF4yK3leMg"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim insertionRange As Range
    Dim insertSeconds(1 To formulaCount) As Double
    Dim insertPhaseTimings(1 To formulaCount) As Variant
    Dim insertNumberingPhaseTimings(1 To formulaCount) As Variant
    Dim resultPath As String
    Dim resultText As String
    Dim updateProfile As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim formulaId As String
    Dim firstFormulaId As String
    Dim lastFormulaId As String
    Dim prependFormulaId As String
    Dim startedAt As Single
    Dim earlyInsertAverage As Double
    Dim lateInsertAverage As Double
    Dim formatUpdateSeconds As Double
    Dim explicitUpdateSeconds As Double
    Dim prependSeconds As Double
    Dim earlyEditSeconds As Double
    Dim lateEditSeconds As Double
    Dim performanceFailures As String
    Dim earlyAppendAverage As Double
    Dim lateAppendAverage As Double
    Dim earlyPictureAverage As Double
    Dim latePictureAverage As Double
    Dim earlyMetadataAverage As Double
    Dim lateMetadataAverage As Double
    Dim earlyNumberingAverage As Double
    Dim lateNumberingAverage As Double
    Dim earlyResolveAverage As Double
    Dim lateResolveAverage As Double
    Dim earlyNumberingSubphase(0 To 5) As Double
    Dim lateNumberingSubphase(0 To 5) As Double
    Dim itemIndex As Long
    Dim phaseIndex As Long

    On Error GoTo RegressionFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-numbering-performance-regression-result.txt"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."

    regressionStage = "sequential-insert-36"
    For itemIndex = 1 To formulaCount
        formulaId = VTRegressionPerformanceFormulaId(itemIndex)
        startedAt = Timer
        Set formulaShape = VTRegressionCreateNumberedImage( _
            testDocument, formulaId, latexBase64, 96!, 28!, _
            insertPhaseTimings(itemIndex), _
            insertNumberingPhaseTimings(itemIndex))
        insertSeconds(itemIndex) = _
            VTTimerElapsedSeconds(startedAt, Timer)
    Next itemIndex
    firstFormulaId = VTRegressionPerformanceFormulaId(1)
    lastFormulaId = VTRegressionPerformanceFormulaId(formulaCount)
    For itemIndex = 6 To 10
        earlyInsertAverage = earlyInsertAverage + insertSeconds(itemIndex)
    Next itemIndex
    earlyInsertAverage = earlyInsertAverage / 5#
    For itemIndex = formulaCount - 4 To formulaCount
        lateInsertAverage = lateInsertAverage + insertSeconds(itemIndex)
    Next itemIndex
    lateInsertAverage = lateInsertAverage / 5#
    For itemIndex = 6 To 10
        earlyAppendAverage = earlyAppendAverage + _
            CDbl(insertPhaseTimings(itemIndex)(0))
        earlyPictureAverage = earlyPictureAverage + _
            CDbl(insertPhaseTimings(itemIndex)(1))
        earlyMetadataAverage = earlyMetadataAverage + _
            CDbl(insertPhaseTimings(itemIndex)(2))
        earlyNumberingAverage = earlyNumberingAverage + _
            CDbl(insertPhaseTimings(itemIndex)(3))
        earlyResolveAverage = earlyResolveAverage + _
            CDbl(insertPhaseTimings(itemIndex)(4))
        For phaseIndex = 0 To 5
            earlyNumberingSubphase(phaseIndex) = _
                earlyNumberingSubphase(phaseIndex) + _
                CDbl(insertNumberingPhaseTimings(itemIndex)(phaseIndex))
        Next phaseIndex
    Next itemIndex
    For itemIndex = formulaCount - 4 To formulaCount
        lateAppendAverage = lateAppendAverage + _
            CDbl(insertPhaseTimings(itemIndex)(0))
        latePictureAverage = latePictureAverage + _
            CDbl(insertPhaseTimings(itemIndex)(1))
        lateMetadataAverage = lateMetadataAverage + _
            CDbl(insertPhaseTimings(itemIndex)(2))
        lateNumberingAverage = lateNumberingAverage + _
            CDbl(insertPhaseTimings(itemIndex)(3))
        lateResolveAverage = lateResolveAverage + _
            CDbl(insertPhaseTimings(itemIndex)(4))
        For phaseIndex = 0 To 5
            lateNumberingSubphase(phaseIndex) = _
                lateNumberingSubphase(phaseIndex) + _
                CDbl(insertNumberingPhaseTimings(itemIndex)(phaseIndex))
        Next phaseIndex
    Next itemIndex
    earlyAppendAverage = earlyAppendAverage / 5#
    lateAppendAverage = lateAppendAverage / 5#
    earlyPictureAverage = earlyPictureAverage / 5#
    latePictureAverage = latePictureAverage / 5#
    earlyMetadataAverage = earlyMetadataAverage / 5#
    lateMetadataAverage = lateMetadataAverage / 5#
    earlyNumberingAverage = earlyNumberingAverage / 5#
    lateNumberingAverage = lateNumberingAverage / 5#
    earlyResolveAverage = earlyResolveAverage / 5#
    lateResolveAverage = lateResolveAverage / 5#
    For phaseIndex = 0 To 5
        earlyNumberingSubphase(phaseIndex) = _
            earlyNumberingSubphase(phaseIndex) / 5#
        lateNumberingSubphase(phaseIndex) = _
            lateNumberingSubphase(phaseIndex) / 5#
    Next phaseIndex
    If lateInsertAverage > earlyInsertAverage * 1.5 + 0.25 Then
        performanceFailures = performanceFailures & _
            "sequential-insert-growth=" & _
            Format$(earlyInsertAverage, "0.000") & "->" & _
            Format$(lateInsertAverage, "0.000") & vbLf
    End If

    VTWordPerformanceStart
    VT_WORD_PERFORMANCE_ACTIVE = True
    regressionStage = "format-update-after-36"
    startedAt = Timer
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_CHAPTER, "."
    VTUpdateEquationNumbersCore testDocument, False
    formatUpdateSeconds = VTTimerElapsedSeconds(startedAt, Timer)
    If formatUpdateSeconds > 10# Then
        performanceFailures = performanceFailures & _
            "format-update-over-10s=" & _
            Format$(formatUpdateSeconds, "0.000") & vbLf
    End If

    updateProfile = "formatProfile" & vbLf & VT_WORD_PERFORMANCE_BUFFER
    VTWordPerformanceStart
    VT_WORD_PERFORMANCE_ACTIVE = True
    regressionStage = "explicit-update-after-36"
    startedAt = Timer
    VTUpdateEquationNumbersCore testDocument, True
    explicitUpdateSeconds = VTTimerElapsedSeconds(startedAt, Timer)
    If explicitUpdateSeconds > 10# Then
        performanceFailures = performanceFailures & _
            "explicit-update-over-10s=" & _
            Format$(explicitUpdateSeconds, "0.000") & vbLf
    End If

    updateProfile = updateProfile & "explicitProfile" & vbLf & _
        VT_WORD_PERFORMANCE_BUFFER
    VT_WORD_PERFORMANCE_ACTIVE = False
    regressionStage = "prepend-before-first"
    Set insertionRange = VTRegressionInsertBlankParagraphBeforeFormula( _
        testDocument, firstFormulaId)
    prependFormulaId = VTRegressionPerformanceFormulaId(999999999)
    startedAt = Timer
    Set formulaShape = VTRegressionCreateNumberedImageAtCaret( _
        testDocument, insertionRange, prependFormulaId, _
        latexBase64, 96!, 28!)
    prependSeconds = VTTimerElapsedSeconds(startedAt, Timer)
    If prependSeconds > 3# Then
        performanceFailures = performanceFailures & _
            "prepend-over-3s=" & Format$(prependSeconds, "0.000") & vbLf
    End If
    VTVerifyNumberedFormulaIntegrity _
        testDocument, prependFormulaId, 1
    VTVerifyNumberedFormulaIntegrity _
        testDocument, firstFormulaId, 2

    regressionStage = "edit-refresh-growth"
    earlyEditSeconds = VTRegressionNumberingRefreshSeconds( _
        testDocument, firstFormulaId, 3)
    lateEditSeconds = VTRegressionNumberingRefreshSeconds( _
        testDocument, lastFormulaId, 3)
    If lateEditSeconds > earlyEditSeconds * 1.5 + 0.25 Then
        performanceFailures = performanceFailures & _
            "edit-refresh-growth=" & _
            Format$(earlyEditSeconds, "0.000") & "->" & _
            Format$(lateEditSeconds, "0.000") & vbLf
    End If

    If Len(performanceFailures) = 0 Then
        resultText = "PASS" & vbLf
    Else
        resultText = "FAIL" & vbLf & _
            "performanceFailures=" & _
            Replace$(performanceFailures, vbLf, ";") & vbLf
    End If
    resultText = resultText & _
        "formulaCount=" & CStr(formulaCount) & vbLf & _
        "insertEarlyAvgSeconds=" & _
            Format$(earlyInsertAverage, "0.000") & vbLf & _
        "insertLateAvgSeconds=" & _
            Format$(lateInsertAverage, "0.000") & vbLf & _
        "insertGrowthRatio=" & _
            Format$(lateInsertAverage / earlyInsertAverage, "0.000") & vbLf
    resultText = resultText & _
        "insertPhaseAppendEarlyLate=" & _
            Format$(earlyAppendAverage, "0.000") & "/" & _
            Format$(lateAppendAverage, "0.000") & vbLf & _
        "insertPhasePictureEarlyLate=" & _
            Format$(earlyPictureAverage, "0.000") & "/" & _
            Format$(latePictureAverage, "0.000") & vbLf & _
        "insertPhaseMetadataEarlyLate=" & _
            Format$(earlyMetadataAverage, "0.000") & "/" & _
            Format$(lateMetadataAverage, "0.000") & vbLf
    resultText = resultText & _
        "insertPhaseNumberingEarlyLate=" & _
            Format$(earlyNumberingAverage, "0.000") & "/" & _
            Format$(lateNumberingAverage, "0.000") & vbLf & _
        "insertPhaseResolveEarlyLate=" & _
            Format$(earlyResolveAverage, "0.000") & "/" & _
            Format$(lateResolveAverage, "0.000") & vbLf
    resultText = resultText & _
        "numberingInitializeEarlyLate=" & _
            Format$(earlyNumberingSubphase(0), "0.000") & "/" & _
            Format$(lateNumberingSubphase(0), "0.000") & vbLf & _
        "numberingTailEarlyLate=" & _
            Format$(earlyNumberingSubphase(1), "0.000") & "/" & _
            Format$(lateNumberingSubphase(1), "0.000") & vbLf
    resultText = resultText & _
        "numberingCenterEarlyLate=" & _
            Format$(earlyNumberingSubphase(2), "0.000") & "/" & _
            Format$(lateNumberingSubphase(2), "0.000") & vbLf & _
        "numberingHelperEarlyLate=" & _
            Format$(earlyNumberingSubphase(3), "0.000") & "/" & _
            Format$(lateNumberingSubphase(3), "0.000") & vbLf
    resultText = resultText & _
        "numberingVisibleEarlyLate=" & _
            Format$(earlyNumberingSubphase(4), "0.000") & "/" & _
            Format$(lateNumberingSubphase(4), "0.000") & vbLf & _
        "numberingFinalizeEarlyLate=" & _
            Format$(earlyNumberingSubphase(5), "0.000") & "/" & _
            Format$(lateNumberingSubphase(5), "0.000") & vbLf
    resultText = resultText & _
        "formatUpdateSeconds=" & _
            Format$(formatUpdateSeconds, "0.000") & vbLf & _
        "explicitUpdateSeconds=" & _
            Format$(explicitUpdateSeconds, "0.000") & vbLf & _
        "prependSeconds=" & Format$(prependSeconds, "0.000") & vbLf
    resultText = resultText & _
        "editEarlyAvgSeconds=" & _
            Format$(earlyEditSeconds, "0.000") & vbLf & _
        "editLateAvgSeconds=" & _
            Format$(lateEditSeconds, "0.000") & vbLf & _
        "editGrowthRatio=" & _
            Format$(lateEditSeconds / earlyEditSeconds, "0.000") & vbLf
    resultText = resultText & updateProfile
    VTWriteTextAtomic resultPath, resultText
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    resultText = "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf & _
        "insertEarlyAvgSeconds=" & _
            Format$(earlyInsertAverage, "0.000") & vbLf & _
        "insertLateAvgSeconds=" & _
            Format$(lateInsertAverage, "0.000") & vbLf & _
        "formatUpdateSeconds=" & _
            Format$(formatUpdateSeconds, "0.000") & vbLf & _
        "explicitUpdateSeconds=" & _
            Format$(explicitUpdateSeconds, "0.000") & vbLf & _
        "prependSeconds=" & Format$(prependSeconds, "0.000") & vbLf & _
        "editEarlyAvgSeconds=" & _
            Format$(earlyEditSeconds, "0.000") & vbLf & _
        "editLateAvgSeconds=" & _
            Format$(lateEditSeconds, "0.000") & vbLf
    VTWriteTextAtomic resultPath, resultText
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word numbering performance regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Private Function VTRegressionCreateNumberedNative( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal latexBase64 As String, _
    ByVal ommlBase64 As String, _
    ByVal nativeDocumentPath As String, _
    Optional ByRef phaseTimings As Variant) As Range

    Dim insertionRange As Range
    Dim equationRange As Range
    Dim numberRange As Range
    Dim encodedMetadata As String
    Dim numberCreated As Boolean
    Dim phaseStartedAt As Single
    Dim appendSeconds As Double
    Dim nativeInsertSeconds As Double
    Dim metadataSeconds As Double
    Dim numberingSeconds As Double
    Dim resolveSeconds As Double
    Dim numberingPhaseTimings As Variant
    Dim nativeInsertPhaseTimings As Variant

    encodedMetadata = VT_METADATA_PREFIX & "e30"
    phaseStartedAt = Timer
    Set insertionRange = VTAppendRegressionParagraph(documentObject)
    appendSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)

    phaseStartedAt = Timer
    Set equationRange = VTInsertNativeEquationAtRange( _
        insertionRange, ommlBase64, nativeDocumentPath, _
        "inline", True, False, nativeInsertPhaseTimings)
    nativeInsertSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)

    phaseStartedAt = Timer
    VTSetWordLatexPayload documentObject, formulaId, latexBase64
    VTSetWordOmmlPayload documentObject, formulaId, ommlBase64
    VTSetWordMetadataPayload documentObject, formulaId, encodedMetadata
    VTSetWordFormulaFormat documentObject, formulaId, "block", True
    metadataSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)

    phaseStartedAt = Timer
    Set numberRange = VTEnsureNativeEquationNumber( _
        equationRange, 48#, formulaId, _
        VTEquationCrossReferenceText(latexBase64), numberCreated, _
        numberingPhaseTimings)
    numberingSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    If Not numberCreated Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The deletion regression native formula did not create a number."
    End If
    phaseStartedAt = Timer
    Set equationRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If equationRange Is Nothing Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The regression native formula identity is missing."
    End If
    If equationRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The regression native formula did not keep its single-paragraph identity."
    End If
    VTSetNativeFormulaBookmark documentObject, equationRange, formulaId
    resolveSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    If Not IsMissing(phaseTimings) Then
        phaseTimings = Array( _
            appendSeconds, nativeInsertSeconds, metadataSeconds, _
            numberingSeconds, resolveSeconds, _
            CDbl(numberingPhaseTimings(0)), _
            CDbl(numberingPhaseTimings(1)), _
            CDbl(numberingPhaseTimings(2)), _
            CDbl(numberingPhaseTimings(3)), _
            CDbl(numberingPhaseTimings(4)), _
            CDbl(numberingPhaseTimings(5)), _
            CDbl(numberingPhaseTimings(6)), _
            CDbl(nativeInsertPhaseTimings(0)), _
            CDbl(nativeInsertPhaseTimings(1)), _
            CDbl(nativeInsertPhaseTimings(2)), _
            CDbl(nativeInsertPhaseTimings(3)), _
            CDbl(nativeInsertPhaseTimings(4)))
    End If
    Set VTRegressionCreateNumberedNative = equationRange.Duplicate
End Function

Private Function VTRegressionInsertBlankParagraphBeforeFormula( _
    ByVal documentObject As Document, _
    ByVal formulaId As String) As Range

    Dim formulaRange As Range
    Dim formulaParagraph As Range
    Dim blankParagraph As Range
    Dim blankStart As Long

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The safe blank-paragraph regression target is invalid."
    End If
    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The neighboring formula is missing before blank-line creation."
    End If
    Set formulaParagraph = _
        VTWordParagraphContainingFormula(formulaRange)
    If formulaParagraph Is Nothing Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "Word could not resolve the neighboring formula paragraph."
    End If
    blankStart = formulaParagraph.Start

    ' Use Word's paragraph operation on the complete formula paragraph. Writing
    ' vbCr into a collapsed Range at OMath.Start is not equivalent to pressing
    ' Enter on a normal line and raises 6193 because Word treats that boundary as
    ' part of the math zone.
    formulaParagraph.InsertParagraphBefore

    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "Creating a blank line removed the neighboring formula identity."
    End If
    Set formulaParagraph = _
        VTWordParagraphContainingFormula(formulaRange)
    If formulaParagraph Is Nothing Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "Word lost the neighboring formula paragraph after blank-line creation."
    End If
    Set blankParagraph = documentObject.Range( _
        Start:=blankStart, End:=blankStart).Paragraphs(1).Range.Duplicate
    If blankParagraph.Start <> blankStart Or _
       blankParagraph.End <> formulaParagraph.Start Or _
       blankParagraph.Information(wdWithInTable) Or _
       blankParagraph.Fields.Count <> 0 Or _
       blankParagraph.InlineShapes.Count <> 0 Or _
       blankParagraph.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(blankParagraph) Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "Word did not create one independent plain paragraph before the formula."
    End If
    VTNormalizePlainWordParagraph blankParagraph
    Set VTRegressionInsertBlankParagraphBeforeFormula = _
        documentObject.Range(Start:=blankStart, End:=blankStart)
End Function

Private Function VTRegressionCreateNumberedNativeAtCaret( _
    ByVal documentObject As Document, _
    ByVal requestedRange As Range, _
    ByVal formulaId As String, _
    ByVal latexBase64 As String, _
    ByVal ommlBase64 As String, _
    ByVal nativeDocumentPath As String) As Range

    Dim insertionRange As Range
    Dim placeholder As InlineShape
    Dim equationRange As Range
    Dim numberRange As Range
    Dim encodedMetadata As String
    Dim operationStage As String
    Dim operationErrorNumber As Long
    Dim operationErrorDescription As String
    Dim nativeStart As Long
    Dim numberCreated As Boolean

    On Error GoTo CreateFailed
    If documentObject Is Nothing Or requestedRange Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The safe native insertion regression target is invalid."
    End If
    encodedMetadata = VT_METADATA_PREFIX & "e30"
    operationStage = "prepare-range"
    Set insertionRange = VTPrepareWordCreateInsertionRange( _
        requestedRange.Duplicate, "block")
    operationStage = "add-placeholder"
    Set placeholder = documentObject.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    placeholder.Width = 1!
    placeholder.Height = 1!

    operationStage = "resolve-placeholder-range"
    Set insertionRange = placeholder.Range.Duplicate
    operationStage = "insert-native"
    Set equationRange = VTInsertNativeEquationAtRange( _
        insertionRange, ommlBase64, nativeDocumentPath, _
        "inline", True, False)
    nativeStart = equationRange.Start
    operationStage = "delete-placeholder"
    placeholder.Delete
    operationStage = "resolve-native"
    Set equationRange = VTResolveNativeEquationRange( _
        documentObject, nativeStart, 16)
    operationStage = "persist-initial-bookmark"
    VTSetNativeFormulaBookmark documentObject, equationRange, formulaId

    operationStage = "write-metadata"
    VTSetWordLatexPayload documentObject, formulaId, latexBase64
    VTSetWordOmmlPayload documentObject, formulaId, ommlBase64
    VTSetWordMetadataPayload documentObject, formulaId, encodedMetadata
    VTSetWordFormulaFormat documentObject, formulaId, "block", True
    operationStage = "number-native"
    Set numberRange = VTEnsureNativeEquationNumber( _
        equationRange, 48#, formulaId, _
        VTEquationCrossReferenceText(latexBase64), numberCreated)
    If Not numberCreated Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The safe native insertion did not create its own Equation number."
    End If

    operationStage = "resolve-numbered-formula"
    Set equationRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If equationRange Is Nothing Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The safe native insertion lost its formula identity."
    End If
    If equationRange.OMaths.Count <> 1 Or _
       equationRange.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The safe native insertion lost its table-free OMath identity."
    End If
    operationStage = "persist-final-bookmark"
    VTSetNativeFormulaBookmark documentObject, equationRange, formulaId
    Set VTRegressionCreateNumberedNativeAtCaret = equationRange.Duplicate
    Exit Function

CreateFailed:
    operationErrorNumber = Err.Number
    operationErrorDescription = Err.Description
    Err.Raise operationErrorNumber, _
        "VisualTeX safe native insertion helper", _
        "VTRegressionCreateNumberedNativeAtCaret/" & operationStage & _
        ": " & operationErrorDescription
End Function

Private Sub VTRegressionAssertExternalNativeSequence( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal expectedOrdinal As Long)

    Dim formulaRange As Range
    Dim formulaParagraph As Range
    Dim helperParagraph As Range
    Dim sequenceField As Field
    Dim visibleNumberField As Field
    Dim candidateField As Field

    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The safe insertion formula cannot be resolved."
    End If
    If formulaRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The safe insertion target is not one native OMath."
    End If
    Set formulaParagraph = _
        VTWordParagraphContainingFormula(formulaRange)
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    Set visibleNumberField = VTNativeEquationArrayReferenceField( _
        formulaRange, formulaId)
    If formulaParagraph Is Nothing Or sequenceField Is Nothing Or _
       visibleNumberField Is Nothing Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The safe insertion SEQ/REF architecture is incomplete."
    End If
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If helperParagraph.Start < formulaParagraph.End Or _
       helperParagraph.OMaths.Count <> 0 Or _
       helperParagraph.InlineShapes.Count <> 0 Or _
       Not VTHelperParagraphOwnsNativeEquationSequence( _
           helperParagraph) Or _
       Not VTNativeEquationNumberIsInsideMath( _
           formulaRange, visibleNumberField) Or _
       VTFirstPositiveIntegerInText(sequenceField.Result.Text) <> _
           expectedOrdinal Or _
       VTFirstPositiveIntegerInText(visibleNumberField.Result.Text) <> _
           expectedOrdinal Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The safe insertion did not keep SEQ outside OMath and REF inside it."
    End If
    For Each candidateField In formulaRange.OMaths(1).Range.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            Err.Raise vbObjectError + 7566, "VisualTeX", _
                "The safe insertion allowed SEQ to be absorbed into OMath."
        End If
    Next candidateField
    VTVerifyNumberedFormulaIntegrity _
        documentObject, formulaId, expectedOrdinal
End Sub

Private Sub VTRegressionAssertPlainCaretAfterNative( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal sentinelText As String)

    Dim formulaRange As Range
    Dim helperParagraph As Range
    Dim caretParagraph As Range
    Dim typedRange As Range
    Dim sequenceField As Field
    Dim textStart As Long

    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    If formulaRange Is Nothing Or sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The safe insertion caret target is incomplete."
    End If
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    VTPlaceCaretAfterDisplayFormula formulaRange, formulaId

    ' Paragraph insertion can expand an already-cached Range. Re-resolve the
    ' helper from its durable VT_N_ Bookmark before comparing final positions.
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The native display caret test lost its Equation SEQ helper."
    End If
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    Set caretParagraph = Selection.Range.Paragraphs(1).Range.Duplicate
    If caretParagraph.Start < helperParagraph.End Or _
       caretParagraph.Start = helperParagraph.Start Or _
       caretParagraph.Information(wdWithInTable) Or _
       caretParagraph.Fields.Count <> 0 Or _
       caretParagraph.InlineShapes.Count <> 0 Or _
       caretParagraph.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The native display caret did not skip its Equation SEQ helper" & _
            " [helper=" & CStr(helperParagraph.Start) & "-" & _
                CStr(helperParagraph.End) & _
            "; caret=" & CStr(caretParagraph.Start) & "-" & _
                CStr(caretParagraph.End) & _
            "; selection=" & CStr(Selection.Start) & "]."
    End If
    textStart = Selection.Start
    Selection.TypeText Text:=sentinelText
    Set typedRange = documentObject.Range( _
        Start:=textStart, End:=Selection.Start)
    If typedRange.Text <> sentinelText Or _
       typedRange.Information(wdWithInTable) Or _
       typedRange.OMaths.Count <> 0 Or _
       typedRange.Fields.Count <> 0 Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "Typing after the native display did not remain ordinary body text."
    End If
End Sub

Private Function VTRegressionPressEnterAfterNumberedDisplay( _
    ByVal documentObject As Document, _
    ByVal formulaId As String) As Range

    Dim formulaRange As Range
    Dim formulaParagraph As Range
    Dim helperParagraph As Range
    Dim blankParagraph As Range
    Dim caretRange As Range
    Dim sequenceField As Field
    Dim caretStart As Long

    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    If formulaRange Is Nothing Or sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The Enter-after-display regression target is incomplete."
    End If
    Set formulaParagraph = _
        VTWordParagraphContainingFormula(formulaRange)
    If formulaParagraph Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The Enter-after-display regression lost the formula paragraph."
    End If

    ' Reproduce the user's exact action: put the caret at the visible paragraph
    ' end and press Return. This deliberately pushes the compact helper away from
    ' its formula, unlike the older regression that inserted a paragraph before a
    ' later formula and therefore missed this bug.
    caretStart = formulaParagraph.End - 1
    Set caretRange = documentObject.Range( _
        Start:=caretStart, End:=caretStart)
    caretRange.Select
    Selection.TypeParagraph
    Set blankParagraph = Selection.Range.Paragraphs(1).Range.Duplicate
    If blankParagraph.Information(wdWithInTable) Or _
       blankParagraph.Fields.Count <> 0 Or _
       blankParagraph.InlineShapes.Count <> 0 Or _
       blankParagraph.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(blankParagraph) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Pressing Return did not create one ordinary blank paragraph."
    End If

    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    Set formulaParagraph = _
        VTWordParagraphContainingFormula(formulaRange)
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If helperParagraph.Start <= formulaParagraph.End Or _
       helperParagraph.Start < blankParagraph.End Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The regression did not reproduce the displaced Equation helper."
    End If
    Set VTRegressionPressEnterAfterNumberedDisplay = _
        documentObject.Range( _
            Start:=blankParagraph.Start, End:=blankParagraph.Start)
End Function

Public Sub VisualTeX_RunWordEnterAfterDisplayNumberingRegression()
    Const nativeFirstId As String = _
        "81818181-8181-4818-8818-818181818181"
    Const nativeSecondId As String = _
        "82828282-8282-4828-8828-828282828282"
    Const imageFirstId As String = _
        "83838383-8383-4838-8838-838383838383"
    Const imageSecondId As String = _
        "84848484-8484-4848-8848-848484848484"
    Const imagePendingSessionId As String = _
        "85858585-8585-4858-8858-858585858585"
    Const fixtureFormulaId As String = _
        "11111111-1111-4111-8111-111111111111"

    Dim testDocument As Document
    Dim nativeRange As Range
    Dim blankRange As Range
    Dim imageShape As InlineShape
    Dim fixtureRoot As String
    Dim nativeDocumentPath As String
    Dim ommlBase64 As String
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    nativeDocumentPath = _
        VTApplicationSupportRoot() & "/NativeDocuments/" & _
        fixtureFormulaId & ".docx"
    resultPath = fixtureRoot & _
        "/word-enter-after-display-numbering-regression-result.txt"
    If Not VTPathFileExists(nativeDocumentPath) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The Enter-after-display native DOCX fixture is missing."
    End If
    ommlBase64 = VTReadText( _
        fixtureRoot & "/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    If Len(ommlBase64) = 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The Enter-after-display OMML fixture is missing."
    End If

    regressionStage = "native-enter-after"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    Set nativeRange = VTRegressionCreateNumberedNative( _
        testDocument, nativeFirstId, "eF8x", ommlBase64, _
        nativeDocumentPath)
    Set blankRange = VTRegressionPressEnterAfterNumberedDisplay( _
        testDocument, nativeFirstId)
    Set nativeRange = VTRegressionCreateNumberedNativeAtCaret( _
        testDocument, blankRange, nativeSecondId, "eF8y", _
        ommlBase64, nativeDocumentPath)
    VTRegressionAssertExternalNativeSequence _
        testDocument, nativeFirstId, 1
    VTRegressionAssertExternalNativeSequence _
        testDocument, nativeSecondId, 2
    If testDocument.Tables.Count <> 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The native Enter-after-display path created a table."
    End If
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing

    regressionStage = "image-create-document"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    regressionStage = "image-create-first"
    Set imageShape = VTRegressionCreateNumberedImage( _
        testDocument, imageFirstId, "eF8x", 96!, 28!)
    regressionStage = "image-press-enter"
    Set blankRange = VTRegressionPressEnterAfterNumberedDisplay( _
        testDocument, imageFirstId)
    regressionStage = "image-pending-bookmark"
    VTAddPendingBookmark blankRange, imagePendingSessionId
    Set blankRange = testDocument.Bookmarks( _
        VTWordBookmarkName(imagePendingSessionId)).Range.Duplicate
    regressionStage = "image-create-second"
    Set imageShape = VTRegressionCreateNumberedImageAtCaret( _
        testDocument, blankRange, imageSecondId, "eF8y", 96!, 28!)
    VTDeletePendingBookmark testDocument, imagePendingSessionId
    regressionStage = "image-verify-first"
    VTVerifyNumberedFormulaIntegrity testDocument, imageFirstId, 1
    regressionStage = "image-verify-second"
    VTVerifyNumberedFormulaIntegrity testDocument, imageSecondId, 2
    regressionStage = "image-verify-structure"
    If testDocument.Tables.Count <> 0 Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "The image Enter-after-display path created a table."
    End If
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing

    VTWriteTextAtomic resultPath, _
        "PASS" & vbLf & _
        "nativeEnterAfter=PASS" & vbLf & _
        "nativeOrdinals=1,2" & vbLf & _
        "imageEnterAfter=PASS" & vbLf & _
        "imageOrdinals=1,2" & vbLf & _
        "tables=0" & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            regressionErrorDescription, vbCr, " "), vbLf, " ") & vbLf
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word Enter-after-display numbering regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordMiddleNativeInsertionRegression()
    Const existingFormulaCount As Long = 24
    Const insertBeforeOrdinal As Long = 21
    Const insertedFormulaId As String = _
        "89898989-8989-4898-8898-898989898989"
    Const fixtureFormulaId As String = _
        "11111111-1111-4111-8111-111111111111"

    Dim testDocument As Document
    Dim formulaRange As Range
    Dim insertionRange As Range
    Dim blankRange As Range
    Dim referenceField As Field
    Dim candidateField As Field
    Dim fixtureRoot As String
    Dim nativeDocumentPath As String
    Dim ommlBase64 As String
    Dim resultPath As String
    Dim formulaId As String
    Dim targetFormulaId As String
    Dim lastFormulaId As String
    Dim sequenceBookmarkName As String
    Dim referenceStart As Long
    Dim nativeItemIndex As Long
    Dim itemIndex As Long
    Dim sequenceCount As Long
    Dim expectedOrdinal As Long
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    nativeDocumentPath = _
        VTApplicationSupportRoot() & "/NativeDocuments/" & _
        fixtureFormulaId & ".docx"
    resultPath = fixtureRoot & _
        "/word-middle-native-insertion-regression-result.txt"
    If Not VTPathFileExists(nativeDocumentPath) Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The middle insertion native DOCX fixture is missing."
    End If
    ommlBase64 = VTReadText( _
        fixtureRoot & "/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    If Len(ommlBase64) = 0 Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The middle insertion OMML fixture is missing."
    End If

    regressionStage = "create-existing-formulas"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    testDocument.Content.Text = "MIDDLE_INSERT_BODY"
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    For itemIndex = 1 To existingFormulaCount
        formulaId = VTRegressionPerformanceFormulaId(2000 + itemIndex)
        Set formulaRange = VTRegressionCreateNumberedNative( _
            testDocument, formulaId, "eF4yK3leMg", ommlBase64, _
            nativeDocumentPath)
    Next itemIndex
    targetFormulaId = _
        VTRegressionPerformanceFormulaId(2000 + insertBeforeOrdinal)
    lastFormulaId = _
        VTRegressionPerformanceFormulaId(2000 + existingFormulaCount)

    regressionStage = "create-cross-reference"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    referenceStart = insertionRange.Start
    nativeItemIndex = VTNativeEquationReferenceItemForFormula( _
        testDocument, lastFormulaId)
    If nativeItemIndex < 1 Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The last existing formula is missing from Word cross-reference items."
    End If
    insertionRange.Select
    Selection.InsertCrossReference _
        ReferenceType:=wdCaptionEquation, _
        ReferenceKind:=wdEntireCaption, _
        ReferenceItem:=nativeItemIndex, _
        InsertAsHyperlink:=True, _
        IncludePosition:=False
    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(lastFormulaId)
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           candidateField.Result.OMaths.Count = 0 And _
           VTEquationFieldStart(candidateField) >= referenceStart And _
           InStr(1, candidateField.Code.Text, sequenceBookmarkName, _
               vbTextCompare) > 0 Then
            Set referenceField = candidateField
            Exit For
        End If
    Next candidateField
    If referenceField Is Nothing Or _
       VTFirstPositiveIntegerInText(referenceField.Result.Text) <> _
           existingFormulaCount Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The pre-insertion cross-reference does not resolve to formula 24."
    End If

    regressionStage = "insert-plain-middle-line"
    Set blankRange = VTRegressionInsertBlankParagraphBeforeFormula( _
        testDocument, targetFormulaId)
    regressionStage = "create-middle-numbered-native"
    Set formulaRange = VTRegressionCreateNumberedNativeAtCaret( _
        testDocument, blankRange, insertedFormulaId, "eF4yK3leMw", _
        ommlBase64, nativeDocumentPath)

    regressionStage = "verify-renumbered-sequence"
    For itemIndex = 1 To existingFormulaCount
        formulaId = VTRegressionPerformanceFormulaId(2000 + itemIndex)
        expectedOrdinal = itemIndex
        If itemIndex >= insertBeforeOrdinal Then
            expectedOrdinal = itemIndex + 1
        End If
        VTRegressionAssertExternalNativeSequence _
            testDocument, formulaId, expectedOrdinal
    Next itemIndex
    VTRegressionAssertExternalNativeSequence _
        testDocument, insertedFormulaId, insertBeforeOrdinal

    regressionStage = "verify-cross-reference"
    Set referenceField = Nothing
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           candidateField.Result.OMaths.Count = 0 And _
           InStr(1, candidateField.Code.Text, sequenceBookmarkName, _
               vbTextCompare) > 0 Then
            Set referenceField = candidateField
            Exit For
        End If
    Next candidateField
    If referenceField Is Nothing Or _
       VTFirstPositiveIntegerInText(referenceField.Result.Text) <> _
           existingFormulaCount + 1 Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The downstream body cross-reference did not flow from 24 to 25."
    End If

    regressionStage = "verify-document-structure"
    For Each candidateField In testDocument.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            sequenceCount = sequenceCount + 1
        End If
    Next candidateField
    If sequenceCount <> existingFormulaCount + 1 Or _
       testDocument.Tables.Count <> 0 Or _
       InStr(1, testDocument.Content.Text, "MIDDLE_INSERT_BODY", _
           vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "Middle insertion changed the document structure or leaked helpers."
    End If

    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    VTWriteTextAtomic resultPath, _
        "PASS" & vbLf & _
        "existingBefore=24" & vbLf & _
        "insertedOrdinal=21" & vbLf & _
        "lastOrdinal=25" & vbLf & _
        "crossReference=25" & vbLf & _
        "tables=0" & vbLf & _
        "externalSeq=PASS" & vbLf & _
        "internalRef=PASS" & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            regressionErrorDescription, vbCr, " "), vbLf, " ") & vbLf
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word middle native insertion regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordSafeNativeInsertionRegression()
    Const bodyEndFormulaId As String = _
        "12121212-1212-4121-8121-121212121212"
    Const bodyGapFormulaId As String = _
        "23232323-2323-4232-8232-232323232323"
    Const bodyGapExistingId As String = _
        "34343434-3434-4343-8343-343434343434"
    Const formulaGapFirstId As String = _
        "45454545-4545-4454-8454-454545454545"
    Const formulaGapInsertedId As String = _
        "56565656-5656-4565-8565-565656565656"
    Const formulaGapSecondId As String = _
        "67676767-6767-4676-8676-676767676767"
    Const fixtureFormulaId As String = _
        "11111111-1111-4111-8111-111111111111"

    Dim testDocument As Document
    Dim formulaRange As Range
    Dim existingRange As Range
    Dim firstRange As Range
    Dim secondRange As Range
    Dim insertionRange As Range
    Dim blankRange As Range
    Dim formulaParagraph As Range
    Dim referenceField As Field
    Dim candidateField As Field
    Dim fixtureRoot As String
    Dim nativeDocumentPath As String
    Dim ommlBase64 As String
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionStep As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim insertionStart As Long
    Dim referenceStart As Long
    Dim nativeItemIndex As Long
    Dim newestReferenceStart As Long

    On Error GoTo RegressionFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    nativeDocumentPath = _
        VTApplicationSupportRoot() & "/NativeDocuments/" & _
        fixtureFormulaId & ".docx"
    resultPath = fixtureRoot & _
        "/word-safe-native-insertion-regression-result.txt"
    If Not VTPathFileExists(nativeDocumentPath) Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The safe insertion native DOCX fixture is missing."
    End If
    ommlBase64 = VTReadText( _
        fixtureRoot & "/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    If Len(ommlBase64) = 0 Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The safe insertion OMML fixture is missing."
    End If

    regressionStage = "body-paragraph-end"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Content.Text = "SAFE_BODY_END"
    insertionStart = testDocument.Content.End - 1
    Set insertionRange = testDocument.Range( _
        Start:=insertionStart, End:=insertionStart)
    Set formulaRange = VTRegressionCreateNumberedNativeAtCaret( _
        testDocument, insertionRange, bodyEndFormulaId, "eF8x", _
        ommlBase64, nativeDocumentPath)
    VTReconcileEquationNumbers testDocument
    VTRegressionAssertExternalNativeSequence _
        testDocument, bodyEndFormulaId, 1
    If InStr(1, testDocument.Content.Text, "SAFE_BODY_END", _
       vbBinaryCompare) = 0 Or testDocument.Tables.Count <> 0 Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "Inserting at a body paragraph end changed the body text or created a table."
    End If
    VTRegressionAssertPlainCaretAfterNative _
        testDocument, bodyEndFormulaId, "SAFE_BODY_END_CONTINUATION"

    regressionStage = "native-entire-caption-reference"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    referenceStart = insertionRange.Start
    nativeItemIndex = VTNativeEquationReferenceItemForFormula( _
        testDocument, bodyEndFormulaId)
    If nativeItemIndex < 1 Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "Word did not expose the external Equation SEQ as a native item."
    End If
    insertionRange.Select
    Selection.InsertCrossReference _
        ReferenceType:=wdCaptionEquation, _
        ReferenceKind:=wdEntireCaption, _
        ReferenceItem:=nativeItemIndex, _
        InsertAsHyperlink:=True, _
        IncludePosition:=False
    newestReferenceStart = -1
    Set referenceField = Nothing
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           candidateField.Result.OMaths.Count = 0 And _
           VTEquationFieldStart(candidateField) >= referenceStart And _
           VTEquationFieldStart(candidateField) > newestReferenceStart Then
            newestReferenceStart = VTEquationFieldStart(candidateField)
            Set referenceField = candidateField
        End If
    Next candidateField
    If referenceField Is Nothing Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "Word did not create an entire-caption Equation reference."
    End If
    If Trim$(referenceField.Result.Text) <> "1" Or _
       referenceField.Result.OMaths.Count <> 0 Or _
       InStr(1, referenceField.Result.Text, "x", _
           vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "Word's entire-caption Equation reference is not the pure number 1."
    End If
    VTRegressionAssertExternalNativeSequence _
        testDocument, bodyEndFormulaId, 1
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing

    regressionStage = "empty-line-between-body-and-formula"
    regressionStep = "create-document"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Content.Text = "SAFE_BODY_ABOVE"
    regressionStep = "create-existing-formula"
    Set existingRange = VTRegressionCreateNumberedNative( _
        testDocument, bodyGapExistingId, "eF8y", ommlBase64, _
        nativeDocumentPath)
    regressionStep = "insert-blank-before-existing"
    Set blankRange = VTRegressionInsertBlankParagraphBeforeFormula( _
        testDocument, bodyGapExistingId)
    regressionStep = "create-formula-in-blank"
    Set formulaRange = VTRegressionCreateNumberedNativeAtCaret( _
        testDocument, blankRange, bodyGapFormulaId, "eF8z", _
        ommlBase64, nativeDocumentPath)
    regressionStep = "reconcile"
    VTReconcileEquationNumbers testDocument
    VTRegressionAssertExternalNativeSequence _
        testDocument, bodyGapFormulaId, 1
    VTRegressionAssertExternalNativeSequence _
        testDocument, bodyGapExistingId, 2
    Set existingRange = VTNumberedFormulaRangeForId( _
        testDocument, bodyGapExistingId)
    If InStr(1, testDocument.Content.Text, "SAFE_BODY_ABOVE", _
       vbBinaryCompare) = 0 Or _
       existingRange Is Nothing Or _
       testDocument.Tables.Count <> 0 Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "Inserting between body text and a formula removed existing content."
    End If
    VTRegressionAssertPlainCaretAfterNative _
        testDocument, bodyGapFormulaId, "SAFE_BODY_GAP_CONTINUATION"
    VTRegressionAssertExternalNativeSequence _
        testDocument, bodyGapExistingId, 2
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing

    regressionStage = "empty-line-between-two-formulas"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    Set firstRange = VTRegressionCreateNumberedNative( _
        testDocument, formulaGapFirstId, "eF8x", ommlBase64, _
        nativeDocumentPath)
    Set secondRange = VTRegressionCreateNumberedNative( _
        testDocument, formulaGapSecondId, "eF8z", ommlBase64, _
        nativeDocumentPath)
    Set blankRange = VTRegressionInsertBlankParagraphBeforeFormula( _
        testDocument, formulaGapSecondId)
    Set formulaRange = VTRegressionCreateNumberedNativeAtCaret( _
        testDocument, blankRange, formulaGapInsertedId, "eF8y", _
        ommlBase64, nativeDocumentPath)
    VTReconcileEquationNumbers testDocument
    VTRegressionAssertExternalNativeSequence _
        testDocument, formulaGapFirstId, 1
    VTRegressionAssertExternalNativeSequence _
        testDocument, formulaGapInsertedId, 2
    VTRegressionAssertExternalNativeSequence _
        testDocument, formulaGapSecondId, 3
    Set firstRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaGapFirstId)
    Set secondRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaGapSecondId)
    If firstRange Is Nothing Or secondRange Is Nothing Or _
       testDocument.Tables.Count <> 0 Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "Inserting between two formulas removed a neighboring formula."
    End If
    VTRegressionAssertPlainCaretAfterNative _
        testDocument, formulaGapInsertedId, _
        "SAFE_FORMULA_GAP_CONTINUATION"
    VTRegressionAssertExternalNativeSequence _
        testDocument, formulaGapFirstId, 1
    VTRegressionAssertExternalNativeSequence _
        testDocument, formulaGapSecondId, 3

    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    VTWriteTextAtomic resultPath, _
        "PASS" & vbLf & _
        "bodyParagraphEnd=PASS" & vbLf & _
        "bodyFormulaGap=PASS" & vbLf & _
        "formulaFormulaGap=PASS" & vbLf & _
        "nativeEntireCaption=1" & vbLf & _
        "tables=0" & vbLf & _
        "externalSeq=PASS" & vbLf & _
        "internalRef=PASS" & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "step=" & regressionStep & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            regressionErrorDescription, vbCr, " "), vbLf, " ") & vbLf
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word safe native insertion regression", _
        regressionStage & "/" & regressionStep & ": " & _
        regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordDeletionReferenceRegression()
    Const imageFormulaId1 As String = _
        "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaa1"
    Const nativeFormulaId2 As String = _
        "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbb2"
    Const nativeFormulaId3 As String = _
        "cccccccc-cccc-4ccc-8ccc-ccccccccccc3"
    Const imageFormulaId4 As String = _
        "dddddddd-dddd-4ddd-8ddd-ddddddddddd4"

    Dim testDocument As Document
    Dim imageFormula1 As InlineShape
    Dim imageFormula4 As InlineShape
    Dim nativeFormula2 As Range
    Dim nativeFormula3 As Range
    Dim layoutTable As Table
    Dim insertionRange As Range
    Dim insertedReference As Range
    Dim formulaIds As Variant
    Dim pickerItems As Variant
    Dim nativeItems As Variant
    Dim candidateField As Field
    Dim fixtureRoot As String
    Dim nativeDocumentPath As String
    Dim ommlBase64 As String
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim liveCount As Long
    Dim sequenceCount As Long
    Dim survivingBodyReferenceCount As Long
    Dim bodyReferenceOneCount As Long
    Dim bodyReferenceTwoCount As Long
    Dim fieldResultText As String

    On Error GoTo RegressionFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    nativeDocumentPath = _
        VTApplicationSupportRoot() & "/NativeDocuments/" & _
        "11111111-1111-4111-8111-111111111111.docx"
    ommlBase64 = VTReadText( _
        fixtureRoot & "/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    resultPath = fixtureRoot & _
        "/word-deletion-reference-regression-result.txt"
    If Not VTPathFileExists(nativeDocumentPath) Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The deletion regression native DOCX fixture is missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."

    regressionStage = "create-four-numbered-formulas"
    Set imageFormula1 = VTRegressionCreateNumberedImage( _
        testDocument, imageFormulaId1, "eF8x", 120!, 34!)
    Set nativeFormula2 = VTRegressionCreateNumberedNative( _
        testDocument, nativeFormulaId2, "eF8y", ommlBase64, _
        nativeDocumentPath)
    Set nativeFormula3 = VTRegressionCreateNumberedNative( _
        testDocument, nativeFormulaId3, "eF8z", ommlBase64, _
        nativeDocumentPath)
    Set imageFormula4 = VTRegressionCreateNumberedImage( _
        testDocument, imageFormulaId4, "eF80", 130!, 38!)
    VTReconcileEquationNumbers testDocument
    formulaIds = VTValidNumberedFormulaIds(testDocument)
    If VTVariantArrayCount(formulaIds) <> 4 Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The deletion regression did not create four live numbered formulas."
    End If

    regressionStage = "insert-surviving-native-references-before-deletion"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set insertedReference = VTInsertEquationNumberReferenceAtRange( _
        insertionRange, 1)
    If insertedReference.Text <> "(1)" Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The pre-deletion native reference to Equation 1 is invalid."
    End If
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set insertedReference = VTInsertEquationNumberReferenceAtRange( _
        insertionRange, 3)
    If insertedReference.Text <> "(3)" Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The pre-deletion native reference to Equation 3 is invalid."
    End If

    regressionStage = "delete-native-and-image-formulas"
    Set nativeFormula2 = VTNumberedFormulaRangeForId( _
        testDocument, nativeFormulaId2)
    nativeFormula2.Delete
    Set insertionRange = VTNumberedFormulaRangeForId( _
        testDocument, imageFormulaId4)
    insertionRange.InlineShapes(1).Delete

    regressionStage = "prune-and-renumber-after-deletion"
    liveCount = VTPruneOrphanedEquationNumberScaffolds(testDocument)
    VTReconcileEquationNumbers testDocument
    If liveCount <> 2 Or testDocument.Tables.Count <> 0 Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "Deleted formulas left orphan number scaffolds" & _
            " [live=" & CStr(liveCount) & _
            "; tables=" & CStr(testDocument.Tables.Count) & "]."
    End If
    formulaIds = VTValidNumberedFormulaIds(testDocument)
    If VTVariantArrayCount(formulaIds) <> 2 Or _
       CStr(formulaIds(1)) <> imageFormulaId1 Or _
       CStr(formulaIds(2)) <> nativeFormulaId3 Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The remaining numbered formulas are not the two live formulas" & _
            " in document order."
    End If
    VTVerifyNumberedFormulaIntegrity _
        testDocument, imageFormulaId1, 1
    VTVerifyNumberedFormulaIntegrity _
        testDocument, nativeFormulaId3, 2

    regressionStage = "verify-deleted-visualtex-artifacts"
    If testDocument.Bookmarks.Exists( _
       VTEquationNumberBookmarkName(nativeFormulaId2)) Or _
       testDocument.Bookmarks.Exists( _
       VTEquationSequenceNumberBookmarkName(nativeFormulaId2)) Or _
       testDocument.Bookmarks.Exists( _
       VTEquationCaptionBookmarkName(nativeFormulaId2)) Or _
       testDocument.Bookmarks.Exists( _
       VTEquationNumberBookmarkName(imageFormulaId4)) Or _
       testDocument.Bookmarks.Exists( _
       VTEquationSequenceNumberBookmarkName(imageFormulaId4)) Or _
       testDocument.Bookmarks.Exists( _
       VTEquationCaptionBookmarkName(imageFormulaId4)) Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "Deleted formulas left VT_R_/VT_N_/VT_C_ Bookmarks."
    End If

    regressionStage = "verify-surviving-native-references-updated"
    For Each candidateField In testDocument.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            sequenceCount = sequenceCount + 1
        ElseIf candidateField.Type = wdFieldRef And _
               Not candidateField.Result.Information(wdWithInTable) And _
               candidateField.Result.Paragraphs(1).Range.InlineShapes.Count = 0 And _
               candidateField.Result.Paragraphs(1).Range.OMaths.Count = 0 Then
            survivingBodyReferenceCount = _
                survivingBodyReferenceCount + 1
            fieldResultText = Trim$(candidateField.Result.Text)
            If VTFirstPositiveIntegerInText(fieldResultText) = 1 Then
                bodyReferenceOneCount = bodyReferenceOneCount + 1
            ElseIf VTFirstPositiveIntegerInText(fieldResultText) = 2 Then
                bodyReferenceTwoCount = bodyReferenceTwoCount + 1
            Else
                Err.Raise vbObjectError + 7554, "VisualTeX", _
                    "A surviving native reference has an invalid result" & _
                    " [code=" & candidateField.Code.Text & _
                    "; result=" & fieldResultText & "]."
            End If
        End If
    Next candidateField
    If sequenceCount <> 2 Or survivingBodyReferenceCount <> 2 Or _
       bodyReferenceOneCount <> 1 Or bodyReferenceTwoCount <> 1 Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The post-deletion native SEQ/REF inventory is invalid" & _
            " [seq=" & CStr(sequenceCount) & _
            "; bodyRef=" & CStr(survivingBodyReferenceCount) & _
            "; one=" & CStr(bodyReferenceOneCount) & _
            "; two=" & CStr(bodyReferenceTwoCount) & "]."
    End If

    regressionStage = "verify-picker-items-match-live-formulas"
    pickerItems = VTEquationNumberCrossReferenceItems(testDocument)
    If VTVariantArrayCount(pickerItems) <> 2 Or _
       InStr(1, CStr(pickerItems(1)), "x_1", vbBinaryCompare) = 0 Or _
       InStr(1, CStr(pickerItems(2)), "x_3", vbBinaryCompare) = 0 Or _
       InStr(1, CStr(pickerItems(1)), "x_2", vbBinaryCompare) > 0 Or _
       InStr(1, CStr(pickerItems(2)), "x_4", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The Equation picker does not match the two live formulas."
    End If
    nativeItems = testDocument.GetCrossReferenceItems(wdCaptionEquation)
    If Not IsArray(nativeItems) Or _
       VTVariantArrayCount(nativeItems) <> 2 Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "Word still exposes orphan native Equation caption items."
    End If

    regressionStage = "insert-fresh-live-references"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set insertedReference = VTInsertEquationNumberReferenceAtRange( _
        insertionRange, 1)
    If insertedReference.Text <> "(1)" Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The fresh reference to the first live formula is invalid."
    End If
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set insertedReference = VTInsertEquationNumberReferenceAtRange( _
        insertionRange, 2)
    If insertedReference.Text <> "(2)" Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The fresh reference to the second live formula is invalid."
    End If

    regressionStage = "reject-broken-native-reference-results"
    bodyReferenceOneCount = 0
    bodyReferenceTwoCount = 0
    survivingBodyReferenceCount = 0
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) And _
           candidateField.Result.Paragraphs(1).Range.InlineShapes.Count = 0 And _
           candidateField.Result.Paragraphs(1).Range.OMaths.Count = 0 Then
            survivingBodyReferenceCount = _
                survivingBodyReferenceCount + 1
            fieldResultText = Trim$(candidateField.Result.Text)
            If VTFirstPositiveIntegerInText(fieldResultText) = 1 Then
                bodyReferenceOneCount = bodyReferenceOneCount + 1
            ElseIf VTFirstPositiveIntegerInText(fieldResultText) = 2 Then
                bodyReferenceTwoCount = bodyReferenceTwoCount + 1
            Else
                Err.Raise vbObjectError + 7554, "VisualTeX", _
                    "A broken native Equation REF remains after cleanup" & _
                    " [code=" & candidateField.Code.Text & _
                    "; result=" & fieldResultText & "]."
            End If
        End If
    Next candidateField
    If survivingBodyReferenceCount <> 4 Or _
       bodyReferenceOneCount <> 2 Or bodyReferenceTwoCount <> 2 Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "Fresh and surviving native references do not match the two" & _
            " live equations."
    End If

    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    VTWriteTextAtomic resultPath, _
        "PASS" & vbLf & _
        "liveFormulas=2" & vbLf & _
        "pickerItems=2" & vbLf & _
        "nativeCaptionItems=2" & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    ' Preserve the failed document for popup/process/range inspection.
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word deletion/reference regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordManagedSequenceInterleaveRegression()
    Const formulaCount As Long = 7
    Const latexBase64 As String = "eF4yK3leMg"

    Dim testDocument As Document
    Dim formulaRange As Range
    Dim sequenceField As Field
    Dim ordinaryField As Field
    Dim restartField As Field
    Dim repeatField As Field
    Dim headingRestartField As Field
    Dim unrelatedSequenceField As Field
    Dim candidateField As Field
    Dim helperParagraph As Range
    Dim insertionRange As Range
    Dim formulaIds As Variant
    Dim fixtureRoot As String
    Dim nativeDocumentPath As String
    Dim ommlBase64 As String
    Dim resultPath As String
    Dim formulaId As String
    Dim lastFormulaId As String
    Dim equationLabelName As String
    Dim legacyErrorDescription As String
    Dim candidateCode As String
    Dim candidateIdentifier As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim helperStart As Long
    Dim legacyErrorNumber As Long
    Dim itemIndex As Long
    Dim nativeSequenceCount As Long
    Dim managedSequenceCount As Long
    Dim ordinarySequenceCount As Long
    Dim headingStart As Long
    Dim restartSwitchPreserved As Boolean
    Dim repeatSwitchPreserved As Boolean
    Dim headingSwitchPreserved As Boolean
    Dim identifierBoundaryPreserved As Boolean

    On Error GoTo RegressionFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    resultPath = fixtureRoot & _
        "/word-managed-sequence-interleave-regression-result.txt"
    nativeDocumentPath = VTApplicationSupportRoot() & _
        "/NativeDocuments/11111111-1111-4111-8111-111111111111.docx"
    ommlBase64 = VTReadText( _
        fixtureRoot & "/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    If Not VTPathFileExists(nativeDocumentPath) Or Len(ommlBase64) = 0 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The managed SEQ interleave fixture is missing."
    End If

    regressionStage = "create-seven-valid-managed-formulas"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    For itemIndex = 1 To formulaCount
        formulaId = VTRegressionPerformanceFormulaId(961000000 + itemIndex)
        Set formulaRange = VTRegressionCreateNumberedNative( _
            testDocument, formulaId, latexBase64, ommlBase64, _
            nativeDocumentPath)
    Next itemIndex
    lastFormulaId = VTRegressionPerformanceFormulaId(961000007)

    ' Insert a normal Word Equation Caption after the seventh visible formula
    ' but before that formula's external VT_N_ helper. Every VisualTeX formula,
    ' VT_N_, VT_R_ and VT_C_ remains valid. Word sees the ordinary Caption as
    ' native item 7 and the managed helper as item 8; the former mixed-anchor
    ' sort instead placed the managed formula at expected item 7.
    regressionStage = "interleave-ordinary-caption-before-seventh-helper"
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, lastFormulaId)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The seventh managed helper is missing."
    End If
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    helperStart = helperParagraph.Start
    helperParagraph.InsertParagraphBefore
    Set helperParagraph = testDocument.Range( _
        Start:=helperStart, End:=helperStart).Paragraphs(1).Range.Duplicate
    VTNormalizePlainWordParagraph helperParagraph
    Set insertionRange = testDocument.Range( _
        Start:=helperStart, End:=helperStart)
    equationLabelName = VTNativeEquationLabelName()
    Set ordinaryField = VTInsertRegisteredEquationCaption( _
        insertionRange, equationLabelName)
    ordinaryField.Update

    regressionStage = "prove-word-eight-visualtex-seven-without-orphan"
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, lastFormulaId)
    sequenceField.Update
    If VTEquationSequenceResultText(sequenceField) <> "8" Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The interleaved native fixture did not produce Word result 8."
    End If
    formulaIds = VTValidNumberedFormulaIds(testDocument)
    If VTVariantArrayCount(formulaIds) <> formulaCount Or _
       VTCountManagedEquationSequences(testDocument) <> formulaCount Or _
       VTPruneDetachedVisualTeXNativeSequenceHelpers(testDocument) <> 0 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The result=8 fixture unexpectedly contains an orphan helper."
    End If

    ' Exercise the exact former expected calculation. The strict validator must
    ' report result=8/expected=7 even though all seven formula identities live.
    On Error Resume Next
    VTApplyEquationSequenceOrdinal _
        sequenceField, equationLabelName, formulaCount
    legacyErrorNumber = Err.Number
    legacyErrorDescription = Err.Description
    Err.Clear
    On Error GoTo RegressionFailed
    If legacyErrorNumber <> vbObjectError + 7549 Or _
       InStr(1, legacyErrorDescription, "result=8", _
           vbTextCompare) = 0 Or _
       InStr(1, legacyErrorDescription, "expected=7", _
           vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The no-orphan fixture did not reproduce result=8/expected=7" & _
            " [error=" & legacyErrorDescription & "]."
    End If

    regressionStage = "reconcile-by-physical-native-sequence-order"
    Set sequenceField = Nothing
    Set ordinaryField = Nothing
    VTReconcileEquationNumbers testDocument, -1, True
    VTVerifyNumberedFormulaIntegrity testDocument, lastFormulaId, 8
    formulaIds = VTValidNumberedFormulaIds(testDocument)
    If VTVariantArrayCount(formulaIds) <> formulaCount Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Reconciliation changed the seven valid formula identities."
    End If

    For Each candidateField In testDocument.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, equationLabelName) Then
            nativeSequenceCount = nativeSequenceCount + 1
            If Len(VTSequenceBookmarkNameForField( _
               testDocument, candidateField)) > 0 Then
                managedSequenceCount = managedSequenceCount + 1
            ElseIf Not VTParagraphHasNumberHelperGeometry( _
                   candidateField.Result.Paragraphs(1).Range.Duplicate) Then
                ordinarySequenceCount = ordinarySequenceCount + 1
            End If
        End If
    Next candidateField
    If nativeSequenceCount <> 8 Or managedSequenceCount <> 7 Or _
       ordinarySequenceCount <> 1 Or _
       VTPruneDetachedVisualTeXNativeSequenceHelpers(testDocument) <> 0 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Reconciliation did not preserve the complete no-orphan inventory."
    End If

    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing

    ' A positional itemIndex also fails for legal user-owned switches. Replay
    ' them without changing their code, then derive the managed field from the
    ' native state: \r 20 -> 20, \c -> 20, Heading 1 + \s 1 displays 1 but
    ' advances the underlying counter to 21, so the next flowing field is 22.
    regressionStage = "replay-native-restart-repeat-heading-switches"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    equationLabelName = VTNativeEquationLabelName()

    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set restartField = testDocument.Fields.Add( _
        Range:=insertionRange, Type:=wdFieldEmpty, _
        Text:="SEQ " & VTEquationSequenceFieldText(equationLabelName) & _
            " \r 20 \* ARABIC", _
        PreserveFormatting:=False)
    restartField.Update
    If VTEquationSequenceResultText(restartField) <> "20" Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Word did not evaluate the ordinary Equation \\r switch as 20."
    End If

    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set repeatField = testDocument.Fields.Add( _
        Range:=insertionRange, Type:=wdFieldEmpty, _
        Text:="SEQ " & VTEquationSequenceFieldText(equationLabelName) & _
            " \c \* ARABIC", _
        PreserveFormatting:=False)
    repeatField.Update
    If VTEquationSequenceResultText(repeatField) <> "20" Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Word did not evaluate the ordinary Equation \\c switch as 20."
    End If

    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    headingStart = insertionRange.Start
    insertionRange.InsertAfter "VisualTeX SEQ regression heading"
    testDocument.Range( _
        Start:=headingStart, End:=headingStart).Paragraphs(1).Style = _
        wdStyleHeading1
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set headingRestartField = testDocument.Fields.Add( _
        Range:=insertionRange, Type:=wdFieldEmpty, _
        Text:="SEQ " & VTEquationSequenceFieldText(equationLabelName) & _
            " \s 1 \* ARABIC", _
        PreserveFormatting:=False)
    headingRestartField.Update
    If VTEquationSequenceResultText(headingRestartField) <> "1" Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Word did not evaluate the ordinary Equation \\s switch as 1."
    End If

    ' A sequence whose identifier merely starts with the localized Equation
    ' label is independent in Word and must not affect VisualTeX expected values.
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set unrelatedSequenceField = testDocument.Fields.Add( _
        Range:=insertionRange, Type:=wdFieldEmpty, _
        Text:="SEQ " & _
            VTEquationSequenceFieldText(equationLabelName & "Appendix") & _
            " \* ARABIC", _
        PreserveFormatting:=False)
    unrelatedSequenceField.Update
    If VTIsNativeEquationSequenceField( _
       unrelatedSequenceField, equationLabelName) Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Equation SEQ discovery matched an identifier prefix."
    End If

    formulaId = VTRegressionPerformanceFormulaId(961100001)
    Set formulaRange = VTRegressionCreateNumberedNative( _
        testDocument, formulaId, latexBase64, ommlBase64, _
        nativeDocumentPath)
    Set restartField = Nothing
    Set repeatField = Nothing
    Set headingRestartField = Nothing
    Set unrelatedSequenceField = Nothing
    VTReconcileEquationNumbers testDocument, -1, True
    VTVerifyNumberedFormulaIntegrity testDocument, formulaId, 22
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldSequence Then
            candidateCode = _
                VTNormalizeEquationFieldCode(candidateField.Code.Text)
            candidateIdentifier = _
                VTSequenceIdentifierFromFieldCode(candidateCode)
            If StrComp(candidateIdentifier, equationLabelName, _
               vbTextCompare) = 0 Then
                If InStr(1, candidateCode, "\r 20", _
                   vbTextCompare) > 0 Then
                    restartSwitchPreserved = _
                        VTEquationSequenceResultText(candidateField) = "20"
                ElseIf InStr(1, candidateCode, "\c", _
                       vbTextCompare) > 0 Then
                    repeatSwitchPreserved = _
                        VTEquationSequenceResultText(candidateField) = "20"
                ElseIf InStr(1, candidateCode, "\s 1", _
                       vbTextCompare) > 0 Then
                    headingSwitchPreserved = _
                        VTEquationSequenceResultText(candidateField) = "1"
                End If
            ElseIf StrComp(candidateIdentifier, _
                   equationLabelName & "Appendix", _
                   vbTextCompare) = 0 Then
                identifierBoundaryPreserved = True
            End If
        End If
    Next candidateField
    If Not restartSwitchPreserved Or _
       Not repeatSwitchPreserved Or _
       Not headingSwitchPreserved Or _
       Not identifierBoundaryPreserved Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "VisualTeX replay changed a user-owned Equation SEQ switch or" & _
            " matched an unrelated identifier prefix."
    End If
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing

    VTWriteTextAtomic resultPath, _
        "PASS" & vbLf & _
        "managedHelpers=7" & vbLf & _
        "validFormulaIds=7" & vbLf & _
        "ordinaryEquationSeq=1" & vbLf & _
        "legacyExpected=7" & vbLf & _
        "wordNativeResult=8" & vbLf & _
        "reconciledExpected=8" & vbLf & _
        "orphanHelpers=0" & vbLf & _
        "restartSwitchResult=20" & vbLf & _
        "repeatSwitchResult=20" & vbLf & _
        "headingSwitchResult=1" & vbLf & _
        "switchFlowManagedResult=22" & vbLf & _
        "identifierBoundary=PASS" & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            regressionErrorDescription, vbCr, " "), vbLf, " ") & vbLf
    If Not testDocument Is Nothing Then
        testDocument.Saved = True
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word managed SEQ interleave regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordOrphanSequenceLifecycleRegression()
    Const latexBase64 As String = "eF4yK3leMg"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim sequenceField As Field
    Dim ordinaryField As Field
    Dim candidateField As Field
    Dim insertionRange As Range
    Dim helperParagraph As Range
    Dim formulaId As String
    Dim firstFormulaId As String
    Dim secondFormulaId As String
    Dim thirdFormulaId As String
    Dim orphanFormulaId As String
    Dim lastFormulaId As String
    Dim sequenceBookmarkName As String
    Dim resultPath As String
    Dim resultText As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim failureErrorNumber As Long
    Dim ordinarySequenceCount As Long
    Dim nativeSequenceCount As Long
    Dim itemIndex As Long
    Dim expectedOrdinal As Long

    On Error GoTo RegressionFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-orphan-sequence-lifecycle-regression-result.txt"

    ' Direct Delete must remove the complete formulaId lifecycle when Update
    ' Numbers discovers that the visible image no longer exists.
    regressionStage = "direct-delete-middle-formula"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    firstFormulaId = VTRegressionPerformanceFormulaId(960100001)
    secondFormulaId = VTRegressionPerformanceFormulaId(960100002)
    thirdFormulaId = VTRegressionPerformanceFormulaId(960100003)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, firstFormulaId, latexBase64, 96!, 28!)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, secondFormulaId, latexBase64, 96!, 28!)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, thirdFormulaId, latexBase64, 96!, 28!)
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, secondFormulaId)
    formulaRange.InlineShapes(1).Delete
    If VTUpdateEquationNumbersCore(testDocument, True) <> 2 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Direct Delete did not leave exactly two numbered formulas."
    End If
    VTVerifyNumberedFormulaIntegrity testDocument, firstFormulaId, 1
    VTVerifyNumberedFormulaIntegrity testDocument, thirdFormulaId, 2
    If testDocument.Bookmarks.Exists( _
           VTEquationNumberBookmarkName(secondFormulaId)) Or _
       testDocument.Bookmarks.Exists( _
           VTEquationSequenceNumberBookmarkName(secondFormulaId)) Or _
       testDocument.Bookmarks.Exists( _
           VTEquationCaptionBookmarkName(secondFormulaId)) Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Direct Delete left formula-specific numbering scaffold."
    End If
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing

    ' Cut follows Word's normal editing path and must converge to the same
    ' lifecycle state without depending on Selection or a cached paragraph.
    regressionStage = "cut-numbered-formula"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    firstFormulaId = VTRegressionPerformanceFormulaId(960200001)
    secondFormulaId = VTRegressionPerformanceFormulaId(960200002)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, firstFormulaId, latexBase64, 96!, 28!)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, secondFormulaId, latexBase64, 96!, 28!)
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, firstFormulaId)
    formulaRange.InlineShapes(1).Range.Cut
    If VTUpdateEquationNumbersCore(testDocument, True) <> 1 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Cut did not leave exactly one numbered formula."
    End If
    VTVerifyNumberedFormulaIntegrity testDocument, secondFormulaId, 1
    If VTCountManagedEquationSequences(testDocument) <> 1 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Cut left a managed Equation SEQ helper."
    End If
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing

    ' Make a real Word Field update fail, then retry immediately. The replay
    ' must rebuild the canonical code/result pair rather than inherit the first
    ' failed attempt's cached Field state.
    regressionStage = "failure-immediate-retry"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    firstFormulaId = VTRegressionPerformanceFormulaId(960300001)
    secondFormulaId = VTRegressionPerformanceFormulaId(960300002)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, firstFormulaId, latexBase64, 96!, 28!)
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, secondFormulaId, latexBase64, 96!, 28!)
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, secondFormulaId)
    CallByName sequenceField, "Locked", VbLet, True
    On Error Resume Next
    VTApplyEquationSequenceOrdinal _
        sequenceField, VTNativeEquationLabelName(), 99
    failureErrorNumber = Err.Number
    Err.Clear
    CallByName sequenceField, "Locked", VbLet, False
    On Error GoTo RegressionFailed
    If failureErrorNumber = 0 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The locked native SEQ did not exercise the strict failure path."
    End If
    Set sequenceField = Nothing
    VTUpdateEquationNumbersCore testDocument, True
    VTVerifyNumberedFormulaIntegrity testDocument, firstFormulaId, 1
    VTVerifyNumberedFormulaIntegrity testDocument, secondFormulaId, 2
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing

    ' Reproduce the reported result=8/expected=7 pollution exactly. Remove the
    ' fourth visible formula and its VT_N_ identity while leaving the private
    ' helper geometry and VT_C_ behind. Reconciliation itself must detect the
    ' orphan before computing ordinals, delete its complete scaffold, take a
    ' fresh Fields snapshot, and replay the remaining seven helpers.
    regressionStage = "orphan-result-eight-expected-seven"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    For itemIndex = 1 To 8
        formulaId = VTRegressionPerformanceFormulaId(960400000 + itemIndex)
        Set formulaShape = VTRegressionCreateNumberedImage( _
            testDocument, formulaId, latexBase64, 96!, 28!)
    Next itemIndex
    orphanFormulaId = VTRegressionPerformanceFormulaId(960400004)
    lastFormulaId = VTRegressionPerformanceFormulaId(960400008)
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, lastFormulaId)
    If VTEquationSequenceResultText(sequenceField) <> "8" Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The eight-formula fixture did not produce native result 8."
    End If

    ' Keep one ordinary Word Equation caption in the same document. It shares
    ' the native sequence name but has neither VisualTeX identity nor the hidden
    ' helper geometry and therefore must survive VisualTeX orphan cleanup.
    regressionStage = "insert-ordinary-word-equation-caption"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set ordinaryField = VTInsertRegisteredEquationCaption( _
        insertionRange, VTNativeEquationLabelName())
    If Len(VTSequenceBookmarkNameForField( _
           testDocument, ordinaryField)) > 0 Or _
       VTParagraphHasNumberHelperGeometry( _
           ordinaryField.Result.Paragraphs(1).Range.Duplicate) Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The ordinary Word Equation caption fixture has VisualTeX identity."
    End If

    regressionStage = "detach-fourth-helper-and-reconcile"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, orphanFormulaId)
    formulaRange.InlineShapes(1).Delete
    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(orphanFormulaId)
    If testDocument.Bookmarks.Exists(sequenceBookmarkName) Then
        testDocument.Bookmarks(sequenceBookmarkName).Delete
    End If
    Set sequenceField = Nothing
    For Each candidateField In testDocument.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            Set helperParagraph = _
                candidateField.Result.Paragraphs(1).Range.Duplicate
            If VTIsDetachedVisualTeXNativeSequenceHelper( _
               helperParagraph) Then
                Set sequenceField = candidateField
                Exit For
            End If
        End If
    Next candidateField
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The detached VisualTeX helper fixture was not recognized."
    End If

    VTReconcileEquationNumbers testDocument, -1, True
    If VTCountManagedEquationSequences(testDocument) <> 7 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Orphan cleanup did not leave seven managed Equation helpers."
    End If
    For itemIndex = 1 To 8
        If itemIndex <> 4 Then
            formulaId = _
                VTRegressionPerformanceFormulaId(960400000 + itemIndex)
            expectedOrdinal = itemIndex
            If itemIndex > 4 Then expectedOrdinal = itemIndex - 1
            VTVerifyNumberedFormulaIntegrity _
                testDocument, formulaId, expectedOrdinal
        End If
    Next itemIndex
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, lastFormulaId)
    If VTEquationSequenceResultText(sequenceField) <> "7" Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The former result 8 did not replay to expected 7."
    End If
    If testDocument.Bookmarks.Exists( _
           VTEquationNumberBookmarkName(orphanFormulaId)) Or _
       testDocument.Bookmarks.Exists( _
           VTEquationSequenceNumberBookmarkName(orphanFormulaId)) Or _
       testDocument.Bookmarks.Exists( _
           VTEquationCaptionBookmarkName(orphanFormulaId)) Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The detached formula left VT_R_/VT_N_/VT_C_ scaffold."
    End If

    ordinarySequenceCount = 0
    nativeSequenceCount = 0
    For Each candidateField In testDocument.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            nativeSequenceCount = nativeSequenceCount + 1
            sequenceBookmarkName = VTSequenceBookmarkNameForField( _
                testDocument, candidateField)
            Set helperParagraph = _
                candidateField.Result.Paragraphs(1).Range.Duplicate
            If Len(sequenceBookmarkName) = 0 And _
               Not VTParagraphHasNumberHelperGeometry( _
                   helperParagraph) Then
                ordinarySequenceCount = ordinarySequenceCount + 1
            End If
        End If
    Next candidateField
    If ordinarySequenceCount <> 1 Or nativeSequenceCount <> 8 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Cleanup deleted or duplicated the ordinary Word Equation caption" & _
            " [ordinary=" & CStr(ordinarySequenceCount) & _
            "; all=" & CStr(nativeSequenceCount) & "]."
    End If

    resultText = "PASS" & vbLf & _
        "directDeleteMiddle=PASS" & vbLf & _
        "cutFormula=PASS" & vbLf & _
        "failureImmediateRetry=PASS" & vbLf & _
        "ordinaryWordEquationSeqPreserved=PASS" & vbLf & _
        "orphanBeforeResult=8" & vbLf & _
        "orphanAfterExpected=7" & vbLf & _
        "managedHelpers=7" & vbLf
    VTWriteTextAtomic resultPath, resultText
    testDocument.Saved = True
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    If Not sequenceField Is Nothing Then _
        CallByName sequenceField, "Locked", VbLet, False
    VTWriteTextAtomic resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            regressionErrorDescription, vbCr, " "), vbLf, " ") & vbLf
    ' Preserve the failed document for popup/process/range inspection.
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word orphan sequence lifecycle regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordReferencePersistenceRegression()
    Const firstFormulaId As String = _
        "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeee1"
    Const referencedFormulaId As String = _
        "ffffffff-ffff-4fff-8fff-fffffffffff2"

    Dim testDocument As Document
    Dim firstFormula As InlineShape
    Dim referencedFormula As InlineShape
    Dim layoutTable As Table
    Dim insertionRange As Range
    Dim insertedReference As Range
    Dim secondReference As Range
    Dim candidateField As Field
    Dim formulaIds As Variant
    Dim fixtureRoot As String
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim bodyReferenceCount As Long
    Dim nativeSequenceCount As Long

    On Error GoTo RegressionFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    resultPath = fixtureRoot & _
        "/word-reference-persistence-regression-result.txt"

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetEquationNumberingFormat _
        testDocument, VT_WORD_NUMBERING_MODE_SEQUENCE, "."

    regressionStage = "create-two-native-caption-formulas"
    Set firstFormula = VTRegressionCreateNumberedImage( _
        testDocument, firstFormulaId, "eF8x", 120!, 34!)
    Set referencedFormula = VTRegressionCreateNumberedImage( _
        testDocument, referencedFormulaId, "eF8y", 130!, 38!)
    VTReconcileEquationNumbers testDocument

    regressionStage = "verify-native-sequence-inventory"
    formulaIds = VTValidNumberedFormulaIds(testDocument)
    If VTVariantArrayCount(formulaIds) <> 2 Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "Word did not preserve exactly two VisualTeX numbered formulas."
    End If
    For Each candidateField In testDocument.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            nativeSequenceCount = nativeSequenceCount + 1
            If InStr(1, candidateField.Code.Text, "\r", _
               vbTextCompare) > 0 Then
                Err.Raise vbObjectError + 7555, "VisualTeX", _
                    "A VisualTeX caption still uses a restarted SEQ field."
            End If
        End If
    Next candidateField
    If nativeSequenceCount <> 2 Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "The two-formula native Equation SEQ inventory is incomplete."
    End If

    regressionStage = "insert-visualtex-native-reference"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set insertedReference = VTInsertEquationNumberReferenceAtRange( _
        insertionRange, 2)
    If insertedReference.Text <> "(2)" Or _
       insertedReference.Fields.Count <> 1 Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "The VisualTeX picker did not insert exactly one reference (2)."
    End If
    VTAssertBodyEquationReferenceVisible _
        testDocument, insertedReference.Fields(1), _
        "initial VisualTeX Equation reference"

    regressionStage = "insert-second-visualtex-native-reference"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set secondReference = VTInsertEquationNumberReferenceAtRange( _
        insertionRange, 2)
    If secondReference.Text <> "(2)" Or _
       secondReference.Fields.Count <> 1 Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "The second VisualTeX native REF is not exactly (2)."
    End If
    VTAssertBodyEquationReferenceVisible _
        testDocument, secondReference.Fields(1), _
        "second VisualTeX Equation reference"

    VTReconcileEquationNumbers testDocument

    regressionStage = "verify-two-live-native-references"
    bodyReferenceCount = 0
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) And _
           candidateField.Result.Paragraphs(1).Range.InlineShapes.Count = 0 And _
           candidateField.Result.Paragraphs(1).Range.OMaths.Count = 0 Then
            bodyReferenceCount = bodyReferenceCount + 1
            If VTFirstPositiveIntegerInText( _
               candidateField.Result.Text) <> 2 Then
                Err.Raise vbObjectError + 7555, "VisualTeX", _
                    "A newly inserted native Equation reference is invalid" & _
                    " [code=" & candidateField.Code.Text & _
                    "; result=" & candidateField.Result.Text & "]."
            End If
            VTAssertBodyEquationReferenceVisible _
                testDocument, candidateField, _
                "new native Equation reference"
        End If
    Next candidateField
    If bodyReferenceCount <> 2 Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "Word did not retain both native Equation references."
    End If

    regressionStage = "delete-preceding-formula-and-renumber"
    Set insertionRange = VTNumberedFormulaRangeForId( _
        testDocument, firstFormulaId)
    insertionRange.InlineShapes(1).Delete
    VTPruneOrphanedEquationNumberScaffolds testDocument
    VTReconcileEquationNumbers testDocument
    formulaIds = VTValidNumberedFormulaIds(testDocument)
    If VTVariantArrayCount(formulaIds) <> 1 Or _
       CStr(formulaIds(1)) <> referencedFormulaId Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "Deleting the first formula did not leave exactly one live" & _
            " VisualTeX numbered formula."
    End If
    bodyReferenceCount = 0
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) And _
           candidateField.Result.Paragraphs(1).Range.InlineShapes.Count = 0 And _
           candidateField.Result.Paragraphs(1).Range.OMaths.Count = 0 Then
            bodyReferenceCount = bodyReferenceCount + 1
            If Trim$(candidateField.Result.Text) <> "1" Then
                Err.Raise vbObjectError + 7555, "VisualTeX", _
                    "A surviving native Equation reference did not renumber" & _
                    " [code=" & candidateField.Code.Text & _
                    "; result=" & candidateField.Result.Text & "]."
            End If
            VTAssertBodyEquationReferenceVisible _
                testDocument, candidateField, _
                "renumbered native Equation reference"
        End If
    Next candidateField
    If bodyReferenceCount <> 2 Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "A surviving native Equation reference disappeared after deletion."
    End If

    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    VTWriteTextAtomic resultPath, _
        "PASS" & vbLf & _
        "initialVisualTeXReferenceA=2" & vbLf & _
        "initialVisualTeXReferenceB=2" & vbLf & _
        "renumberedVisualTeXReferenceA=1" & vbLf & _
        "renumberedVisualTeXReferenceB=1" & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            regressionErrorDescription, vbCr, " "), vbLf, " ") & vbLf
    ' Preserve the failed document for popup/process/range inspection.
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word native reference persistence regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_RunWordDisplayStrategyProbe()
    Dim fixtureRoot As String
    Dim nativeRoot As String
    Dim resultPath As String
    Dim report As String
    Dim fixtureNames As Variant
    Dim ommlFiles As Variant
    Dim documentIds As Variant
    Dim strategies As Variant
    Dim fixtureIndex As Long
    Dim strategyIndex As Long
    Dim ommlPath As String
    Dim nativeDocumentPath As String
    Dim ommlBase64 As String
    Dim probeErrorNumber As Long
    Dim probeErrorDescription As String

    On Error GoTo ProbeFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    nativeRoot = VTApplicationSupportRoot() & "/NativeDocuments"
    resultPath = fixtureRoot & "/word-display-strategy-probe-result.txt"
    fixtureNames = Array("fraction", "integral", "sum_fraction")
    ' The OMML string is used only for payload validation in the insertion
    ' helper; the actual native structure comes from each fixture DOCX below.
    ommlFiles = Array( _
        "word-native-regression-omml.txt", _
        "word-native-regression-omml.txt", _
        "word-native-regression-omml.txt")
    documentIds = Array( _
        "11111111-1111-4111-8111-111111111111", _
        "66666666-6666-4666-8666-666666666666", _
        "77777777-7777-4777-8777-777777777777")
    strategies = Array( _
        "inline-baseline", _
        "formatted-paragraph", _
        "formatted-paragraph-compact-tail", _
        "copy-paste-paragraph", _
        "paste-original-paragraph", _
        "paste-rtf-paragraph", _
        "cut-paste-paragraph", _
        "copy-paste-equation", _
        "linearize-display-before-build", _
        "linearize-build-before-display")

    report = _
        "state=RUNNING" & vbLf & _
        "probeVersion=1" & vbLf & _
        "columns=fixture|strategy|status|stage|error|sourceType|" & _
        "sourceMathPara|cellType|cellMathPara|cellMaths|cellParagraphs|" & _
        "formulaAdvance|anchorSpan|fontSize|linearLength|description" & vbLf
    VTWriteTextAtomic resultPath, report

    For fixtureIndex = LBound(fixtureNames) To UBound(fixtureNames)
        ommlPath = fixtureRoot & "/" & CStr(ommlFiles(fixtureIndex))
        nativeDocumentPath = nativeRoot & "/" & _
            CStr(documentIds(fixtureIndex)) & ".docx"
        If Not VTPathFileExists(ommlPath) Or _
           Not VTPathFileExists(nativeDocumentPath) Then
            report = report & _
                CStr(fixtureNames(fixtureIndex)) & _
                "|all|MISSING|fixture-check|0|||||||||||" & _
                "OMML or DOCX fixture is missing" & vbLf
            VTWriteTextAtomic resultPath, report
        Else
            ommlBase64 = VTReadText( _
                ommlPath, _
                VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
            For strategyIndex = LBound(strategies) To UBound(strategies)
                report = report & VTProbeOneDisplayStrategy( _
                    CStr(fixtureNames(fixtureIndex)), _
                    CStr(strategies(strategyIndex)), _
                    ommlBase64, _
                    nativeDocumentPath) & vbLf
                VTWriteTextAtomic resultPath, report
            Next strategyIndex
        End If
    Next fixtureIndex

    report = report & "state=COMPLETE" & vbLf
    VTWriteTextAtomic resultPath, report
    Exit Sub

ProbeFailed:
    probeErrorNumber = Err.Number
    probeErrorDescription = Err.Description
    On Error Resume Next
    report = report & _
        "state=FATAL" & vbLf & _
        "errorNumber=" & CStr(probeErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(probeErrorDescription, vbCr, " "), vbLf, " ") & _
            vbLf
    VTWriteTextAtomic resultPath, report
    On Error GoTo 0
End Sub

Private Function VTProbeOneDisplayStrategy( _
    ByVal fixtureName As String, _
    ByVal strategyName As String, _
    ByVal ommlBase64 As String, _
    ByVal nativeDocumentPath As String) As String

    Dim probeDocument As Document
    Dim insertionRange As Range
    Dim sourceRange As Range
    Dim sourceParagraph As Range
    Dim centerRange As Range
    Dim formulaRange As Range
    Dim formulaParagraph As Range
    Dim beforeAnchorRange As Range
    Dim afterTableParagraph As Range
    Dim layoutTable As Table
    Dim sourceEquation As OMath
    Dim centerEquation As OMath
    Dim sourceParagraphStart As Long
    Dim beforeAnchorStart As Long
    Dim centerStart As Long
    Dim sourceXml As String
    Dim cellXml As String
    Dim linearText As String
    Dim sourceType As Long
    Dim cellType As Long
    Dim sourceMathParaCount As Long
    Dim cellMathParaCount As Long
    Dim cellMathCount As Long
    Dim cellParagraphCount As Long
    Dim formulaY As Single
    Dim beforeAnchorY As Single
    Dim afterTableY As Single
    Dim formulaAdvance As Single
    Dim anchorSpan As Single
    Dim fontSize As Single
    Dim probeStage As String
    Dim probeErrorNumber As Long
    Dim probeErrorDescription As String

    On Error GoTo StrategyFailed
    probeStage = "create-document"
    Set probeDocument = Documents.Add(Visible:=True)
    probeDocument.ActiveWindow.View.Type = wdPrintView
    probeDocument.Activate

    probeStage = "insert-source-equation"
    Set insertionRange = probeDocument.Range(Start:=0, End:=0)
    Set sourceRange = VTInsertNativeEquationAtRange( _
        insertionRange, ommlBase64, nativeDocumentPath, _
        "inline", True, False)
    If sourceRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7550, "VisualTeX", _
            "The strategy probe did not create one source OMath."
    End If
    Set sourceEquation = sourceRange.OMaths(1)
    sourceEquation.Type = wdOMathInline
    sourceEquation.Range.Font.Position = 0
    Set sourceParagraph = _
        sourceEquation.Range.Paragraphs(1).Range.Duplicate
    sourceParagraphStart = sourceParagraph.Start
    sourceType = sourceEquation.Type
    sourceXml = VTProbeRangeWordOpenXml(sourceParagraph)
    sourceMathParaCount = _
        VTProbeSubstringCount(sourceXml, "<m:oMathPara")
    If VTProbeParagraphHasOutsideContent( _
       sourceParagraph, sourceEquation.Range) Then
        Err.Raise vbObjectError + 7550, "VisualTeX", _
            "The strategy probe source OMath is not alone in its paragraph."
    End If

    probeStage = "create-before-anchor"
    Set insertionRange = VTAppendRegressionParagraph(probeDocument)
    beforeAnchorStart = insertionRange.Start
    insertionRange.Text = "P"
    Set beforeAnchorRange = probeDocument.Range( _
        Start:=beforeAnchorStart, End:=beforeAnchorStart + 1)
    probeDocument.Bookmarks.Add _
        Name:="VTProbeBeforeTable", Range:=beforeAnchorRange

    probeStage = "create-empty-table"
    Set insertionRange = VTAppendRegressionParagraph(probeDocument)
    Set layoutTable = probeDocument.Tables.Add( _
        Range:=insertionRange, NumRows:=1, NumColumns:=3)
    VTConfigureNumberedDisplayTable layoutTable

    If strategyName = "inline-baseline" Then
        probeStage = "insert-inline-baseline"
        Set sourceParagraph = probeDocument.Range( _
            Start:=sourceParagraphStart, _
            End:=sourceParagraphStart).Paragraphs(1).Range.Duplicate
        Set sourceEquation = sourceParagraph.OMaths(1)
        Set centerRange = layoutTable.Cell(1, 2).Range.Duplicate
        centerRange.End = centerRange.End - 1
        centerRange.FormattedText = sourceEquation.Range.FormattedText
    Else
        probeStage = "promote-source-display"
        Set sourceParagraph = probeDocument.Range( _
            Start:=sourceParagraphStart, _
            End:=sourceParagraphStart).Paragraphs(1).Range.Duplicate
        Set sourceEquation = sourceParagraph.OMaths(1)
        sourceEquation.Type = wdOMathDisplay
        sourceEquation.Justification = wdOMathJcCenter
        sourceEquation.Range.Font.Position = 0
        sourceEquation.Range.Font.Size = _
            VTPreferredNativeDisplayFontSize(sourceEquation.Range)
        sourceEquation.BuildUp
        Set sourceParagraph = probeDocument.Range( _
            Start:=sourceParagraphStart, _
            End:=sourceParagraphStart).Paragraphs(1).Range.Duplicate
        If sourceParagraph.OMaths.Count <> 1 Then
            Err.Raise vbObjectError + 7550, "VisualTeX", _
                "The strategy probe lost its source display OMath."
        End If
        Set sourceEquation = sourceParagraph.OMaths(1)
        sourceType = sourceEquation.Type
        sourceXml = VTProbeRangeWordOpenXml(sourceParagraph)
        sourceMathParaCount = _
            VTProbeSubstringCount(sourceXml, "<m:oMathPara")

        probeStage = "apply-strategy"
        Set centerRange = layoutTable.Cell(1, 2).Range.Duplicate
        centerRange.End = centerRange.End - 1
        Select Case strategyName
            Case "formatted-paragraph"
                centerRange.FormattedText = sourceParagraph.FormattedText
            Case "formatted-paragraph-compact-tail"
                centerRange.FormattedText = sourceParagraph.FormattedText
                VTCompactNativeDisplayCellTail layoutTable
            Case "copy-paste-paragraph"
                sourceParagraph.Copy
                centerRange.Paste
            Case "paste-original-paragraph"
                sourceParagraph.Copy
                centerRange.PasteAndFormat wdFormatOriginalFormatting
            Case "paste-rtf-paragraph"
                sourceParagraph.Copy
                centerRange.PasteSpecial DataType:=wdPasteRTF
            Case "cut-paste-paragraph"
                sourceParagraph.Cut
                centerRange.Paste
            Case "copy-paste-equation"
                sourceEquation.Range.Copy
                centerRange.Paste
            Case "linearize-display-before-build", _
                 "linearize-build-before-display"
                sourceEquation.Linearize
                Set sourceParagraph = probeDocument.Range( _
                    Start:=sourceParagraphStart, _
                    End:=sourceParagraphStart).Paragraphs(1).Range.Duplicate
                linearText = VTProbeParagraphText(sourceParagraph)
                If Len(linearText) = 0 Then
                    Err.Raise vbObjectError + 7550, "VisualTeX", _
                        "Word returned an empty linearized equation."
                End If
                centerStart = layoutTable.Cell(1, 2).Range.Start
                Set centerRange = layoutTable.Cell(1, 2).Range.Duplicate
                centerRange.End = centerRange.End - 1
                centerRange.Text = linearText
                Set formulaRange = probeDocument.Range( _
                    Start:=centerStart, _
                    End:=centerStart + Len(linearText))
                Set formulaRange = probeDocument.OMaths.Add(formulaRange)
                If formulaRange.OMaths.Count <> 1 Then
                    Err.Raise vbObjectError + 7550, "VisualTeX", _
                        "Word did not rebuild one OMath from linear text."
                End If
                Set centerEquation = formulaRange.OMaths(1)
                If strategyName = _
                   "linearize-display-before-build" Then
                    centerEquation.Type = wdOMathDisplay
                    centerEquation.Justification = wdOMathJcCenter
                    centerEquation.BuildUp
                Else
                    centerEquation.BuildUp
                    centerEquation.Type = wdOMathDisplay
                    centerEquation.Justification = wdOMathJcCenter
                    centerEquation.BuildUp
                End If
            Case Else
                Err.Raise vbObjectError + 7550, "VisualTeX", _
                    "The display strategy probe name is invalid."
        End Select
    End If
    DoEvents

    probeStage = "finalize-center-equation"
    cellMathCount = layoutTable.Cell(1, 2).Range.OMaths.Count
    If cellMathCount <> 1 Then
        Err.Raise vbObjectError + 7550, "VisualTeX", _
            "The strategy did not leave exactly one center-cell OMath."
    End If
    Set centerEquation = layoutTable.Cell(1, 2).Range.OMaths(1)
    If strategyName = "inline-baseline" Then
        centerEquation.Type = wdOMathInline
    Else
        centerEquation.Type = wdOMathDisplay
        centerEquation.Justification = wdOMathJcCenter
    End If
    centerEquation.BuildUp
    Set centerEquation = layoutTable.Cell(1, 2).Range.OMaths(1)
    centerEquation.Range.Font.Position = 0
    centerEquation.Range.Font.Size = _
        VTPreferredNativeDisplayFontSize(layoutTable.Cell(1, 2).Range)
    layoutTable.Cell(1, 2).Range.ParagraphFormat.Alignment = _
        wdAlignParagraphCenter

    probeStage = "inspect-cell"
    cellType = centerEquation.Type
    fontSize = centerEquation.Range.Font.Size
    cellParagraphCount = _
        layoutTable.Cell(1, 2).Range.Paragraphs.Count
    cellXml = VTProbeRangeWordOpenXml( _
        layoutTable.Cell(1, 2).Range)
    cellMathParaCount = _
        VTProbeSubstringCount(cellXml, "<m:oMathPara")

    probeStage = "measure-table"
    probeDocument.Repaginate
    Set formulaParagraph = _
        layoutTable.Cell(1, 2).Range.Paragraphs(1).Range.Duplicate
    Set afterTableParagraph = probeDocument.Range( _
        Start:=layoutTable.Range.End, _
        End:=layoutTable.Range.End).Paragraphs(1).Range.Duplicate
    formulaY = CSng(formulaParagraph.Information( _
        wdVerticalPositionRelativeToPage))
    afterTableY = CSng(afterTableParagraph.Information( _
        wdVerticalPositionRelativeToPage))
    formulaAdvance = afterTableY - formulaY
    If Not probeDocument.Bookmarks.Exists("VTProbeBeforeTable") Then
        Err.Raise vbObjectError + 7550, "VisualTeX", _
            "The fixed before-table measurement Bookmark disappeared."
    End If
    Set beforeAnchorRange = probeDocument.Bookmarks( _
        "VTProbeBeforeTable").Range.Paragraphs(1).Range.Duplicate
    beforeAnchorY = CSng(beforeAnchorRange.Information( _
        wdVerticalPositionRelativeToPage))
    anchorSpan = afterTableY - beforeAnchorY

    VTProbeOneDisplayStrategy = _
        fixtureName & "|" & strategyName & _
        "|OK|complete|0|" & CStr(sourceType) & "|" & _
        CStr(sourceMathParaCount) & "|" & CStr(cellType) & "|" & _
        CStr(cellMathParaCount) & "|" & CStr(cellMathCount) & "|" & _
        CStr(cellParagraphCount) & "|" & CStr(formulaAdvance) & "|" & _
        CStr(anchorSpan) & "|" & CStr(fontSize) & "|" & _
        CStr(Len(linearText)) & "|"

    probeDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set probeDocument = Nothing
    Exit Function

StrategyFailed:
    probeErrorNumber = Err.Number
    probeErrorDescription = Err.Description
    On Error Resume Next
    If Not probeDocument Is Nothing Then
        probeDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    VTProbeOneDisplayStrategy = _
        fixtureName & "|" & strategyName & "|FAIL|" & _
        probeStage & "|" & CStr(probeErrorNumber) & _
        "|||||||||||" & _
        Replace$(Replace$(probeErrorDescription, vbCr, " "), vbLf, " ")
End Function

Private Function VTProbeParagraphHasOutsideContent( _
    ByVal paragraphRange As Range, _
    ByVal equationRange As Range) As Boolean

    Dim beforeRange As Range
    Dim afterRange As Range

    If paragraphRange Is Nothing Or equationRange Is Nothing Then
        VTProbeParagraphHasOutsideContent = True
        Exit Function
    End If
    Set beforeRange = paragraphRange.Duplicate
    beforeRange.End = equationRange.Start
    Set afterRange = paragraphRange.Duplicate
    afterRange.Start = equationRange.End
    VTProbeParagraphHasOutsideContent = _
        VTWordRangeHasMeaningfulText(beforeRange) Or _
        VTWordRangeHasMeaningfulText(afterRange)
End Function

Private Function VTProbeRangeWordOpenXml( _
    ByVal targetRange As Range) As String

    On Error Resume Next
    VTProbeRangeWordOpenXml = targetRange.WordOpenXML
    On Error GoTo 0
End Function

Private Function VTProbeSubstringCount( _
    ByVal value As String, _
    ByVal fragment As String) As Long

    Dim searchStart As Long
    Dim matchStart As Long

    If Len(value) = 0 Or Len(fragment) = 0 Then Exit Function
    searchStart = 1
    Do
        matchStart = InStr(searchStart, value, fragment, vbBinaryCompare)
        If matchStart = 0 Then Exit Do
        VTProbeSubstringCount = VTProbeSubstringCount + 1
        searchStart = matchStart + Len(fragment)
    Loop
End Function

Private Function VTProbeParagraphText( _
    ByVal paragraphRange As Range) As String

    Dim value As String
    Dim trailingCharacter As String

    If paragraphRange Is Nothing Then Exit Function
    value = paragraphRange.Text
    Do While Len(value) > 0
        trailingCharacter = Right$(value, 1)
        If trailingCharacter = vbCr Or _
           trailingCharacter = vbLf Or _
           trailingCharacter = Chr$(7) Then
            value = Left$(value, Len(value) - 1)
        Else
            Exit Do
        End If
    Loop
    VTProbeParagraphText = value
End Function

Public Sub VisualTeX_RunNativeToImageFastPerformanceRegression()
    Const formulaId As String = _
        "0274750e-c676-4f7b-8c04-a9d9ef7336a0"

    Dim testDocument As Document
    Dim insertionRange As Range
    Dim equationRange As Range
    Dim nativeMath As OMath
    Dim nativeBookmark As Bookmark
    Dim resultPath As String
    Dim phaseStartedAt As Single
    Dim elapsedSeconds As Double
    Dim performancePhases As Variant
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-native-to-image-fast-performance-result.txt"
    On Error GoTo RegressionFailed

    If Not VTPathFileExists(VTNativeWordDocumentPath(formulaId)) Or _
       Not VTPathFileExists(VTImageCacheDocumentPath(formulaId)) Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The native-to-image performance cache fixture is missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.Activate
    testDocument.ActiveWindow.View.Type = wdPrintView
    Set insertionRange = testDocument.Content.Duplicate
    insertionRange.Collapse wdCollapseStart
    Set equationRange = VTInsertCachedNativeEquationFast( _
        insertionRange, VTNativeWordDocumentPath(formulaId), 14#)
    Set equationRange = VTPromoteNativeEquationToDisplay(equationRange)

    VTSetWordLatexPayload _
        testDocument, formulaId, VTBase64UrlEncodeUtf8("x")
    VTSetWordMetadataPayload _
        testDocument, formulaId, VT_METADATA_PREFIX & "e30"
    VTSetWordFormulaFormat testDocument, formulaId, "block", False
    VTSetWordImageScaleState _
        testDocument, formulaId, 14#, _
        800300# / 12700#, 422283# / 12700#, 0#, 14#
    VTSetNativeFormulaBookmark testDocument, equationRange, formulaId
    Set nativeMath = VTNativeMathForBookmark( _
        testDocument.Bookmarks(VTNativeFormulaBookmarkName(formulaId)))
    If nativeMath Is Nothing Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The native-to-image performance fixture lost its OMath."
    End If
    VTSetWordNativeSignature testDocument, formulaId, nativeMath
    Set nativeBookmark = _
        testDocument.Bookmarks(VTNativeFormulaBookmarkName(formulaId))
    nativeBookmark.Range.Select

    phaseStartedAt = Timer
    If Not VTWordConvertNativeBookmarkToImageFast( _
       nativeBookmark, performancePhases) Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The native-to-image performance fixture did not use the fast path."
    End If
    elapsedSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)

    If testDocument.InlineShapes.Count <> 1 Or _
       testDocument.OMaths.Count <> 0 Or _
       testDocument.Tables.Count <> 0 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The native-to-image fast path did not produce one canonical image formula."
    End If

    VTWriteTextAtomic resultPath, _
        "PASS" & vbLf & _
        "elapsedMs=" & CStr(CLng(elapsedSeconds * 1000#)) & vbLf & _
        "resolveMs=" & CStr(CLng(CDbl(performancePhases(0)) * 1000#)) & vbLf & _
        "metadataMs=" & CStr(CLng(CDbl(performancePhases(1)) * 1000#)) & vbLf & _
        "cacheMs=" & CStr(CLng(CDbl(performancePhases(2)) * 1000#)) & vbLf & _
        "transactionMs=" & CStr(CLng(CDbl(performancePhases(3)) * 1000#)) & vbLf & _
        "stagingMs=" & CStr(CLng(CDbl(performancePhases(4)) * 1000#)) & vbLf & _
        "finalizeMs=" & CStr(CLng(CDbl(performancePhases(5)) * 1000#)) & vbLf
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic resultPath, _
        "FAIL" & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX native-to-image performance regression", _
        regressionErrorDescription
End Sub

Public Sub VisualTeX_SetupNumberedConversionRibbonRegression()
    Const formulaId As String = _
        "0274750e-c676-4f7b-8c04-a9d9ef7336a0"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim ommlBase64 As String

    ommlBase64 = VTReadText( _
        VTApplicationSupportRoot() & _
            "/Tests/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    If Len(ommlBase64) = 0 Or _
       Not VTPathFileExists(VTNativeWordDocumentPath(formulaId)) Or _
       Not VTPathFileExists(VTImageCacheDocumentPath(formulaId)) Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The numbered Ribbon conversion fixture is missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.Activate
    testDocument.ActiveWindow.View.Type = wdPrintView
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_MODE_VARIABLE, _
        VT_WORD_NUMBERING_MODE_SEQUENCE
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_SEPARATOR_VARIABLE, "."
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, formulaId, VTBase64UrlEncodeUtf8("x"), _
        63.01575!, 33.25063!)
    VTSetWordOmmlPayload testDocument, formulaId, ommlBase64
    VTSetWordImageScaleState _
        testDocument, formulaId, 14#, 63.01575#, 33.25063#, 0#, 14#
    formulaShape.Select
End Sub

Public Sub VisualTeX_InspectNumberedConversionRibbonRegression()
    Const formulaId As String = _
        "0274750e-c676-4f7b-8c04-a9d9ef7336a0"

    Dim documentObject As Document
    Dim formulaRange As Range
    Dim nativeMath As OMath
    Dim visibleNumberField As Field
    Dim formulaKind As String
    Dim hasVisibleRef As Boolean
    Dim hasArrayMarker As Boolean
    Dim resultPath As String

    If Documents.Count = 0 Then Exit Sub
    Set documentObject = ActiveDocument
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-numbered-conversion-ribbon-result.txt"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If formulaRange Is Nothing Then
        VTWriteTextAtomic resultPath, "FAIL" & vbLf & "reason=missing-formula" & vbLf
        Exit Sub
    End If
    If formulaRange.InlineShapes.Count = 1 And _
       formulaRange.OMaths.Count = 0 Then
        formulaKind = "image"
        Set visibleNumberField = VTImageEquationReferenceField( _
            formulaRange, formulaId)
        hasVisibleRef = Not visibleNumberField Is Nothing
    ElseIf formulaRange.InlineShapes.Count = 0 And _
           formulaRange.OMaths.Count = 1 Then
        formulaKind = "omml"
        Set nativeMath = formulaRange.OMaths(1)
        hasArrayMarker = Not _
            VTNativeEquationArrayMarkerRange(nativeMath.Range.Duplicate) Is Nothing
        Set visibleNumberField = VTNativeEquationArrayReferenceField( _
            nativeMath.Range.Duplicate, formulaId)
        If Not visibleNumberField Is Nothing Then
            hasVisibleRef = VTNativeEquationNumberIsInsideMath( _
                nativeMath.Range.Duplicate, visibleNumberField)
        End If
    Else
        formulaKind = "ambiguous"
    End If
    VTWriteTextAtomic resultPath, _
        "kind=" & formulaKind & vbLf & _
        "number=" & VTEquationNumberTextForFormula( _
            documentObject, formulaId) & vbLf & _
        "tables=" & CStr(documentObject.Tables.Count) & vbLf & _
        "sequences=" & CStr(VTCountManagedEquationSequences( _
            documentObject)) & vbLf & _
        "visibleRef=" & CStr(hasVisibleRef) & vbLf & _
        "arrayMarker=" & CStr(hasArrayMarker) & vbLf
End Sub

Private Function VTRegressionCreateGlobalFormatImage( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal encodedMetadata As String, _
    ByVal latexText As String, _
    ByVal displayMode As String, _
    ByVal numbered As Boolean, _
    ByVal widthPoints As Single, _
    ByVal heightPoints As Single) As InlineShape

    Dim insertionRange As Range
    Dim numberRange As Range
    Dim formulaShape As InlineShape
    Dim latexBase64 As String

    If documentObject Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Or _
       Not VTIsEncodedMetadata(encodedMetadata) Or _
       Len(latexText) = 0 Then
        Err.Raise vbObjectError + 7600, "VisualTeX regression", _
            "The global format conversion image fixture is invalid."
    End If
    If displayMode <> "inline" And displayMode <> "block" Then
        Err.Raise vbObjectError + 7600, "VisualTeX regression", _
            "The global format conversion display mode is invalid."
    End If
    If numbered And displayMode <> "block" Then
        Err.Raise vbObjectError + 7600, "VisualTeX regression", _
            "Only block formulas can be numbered in the global conversion fixture."
    End If

    latexBase64 = VTBase64UrlEncodeUtf8(latexText)
    Set insertionRange = VTAppendRegressionParagraph(documentObject)
    Set formulaShape = documentObject.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    formulaShape.LockAspectRatio = msoFalse
    formulaShape.Width = widthPoints
    formulaShape.Height = heightPoints
    formulaShape.LockAspectRatio = msoTrue
    formulaShape.AlternativeText = encodedMetadata
    formulaShape.Title = VTFormulaReference( _
        formulaId, displayMode, numbered)
    formulaShape.Range.Font.Size = 14!
    formulaShape.Range.Font.Position = 0

    VTSetWordLatexPayload documentObject, formulaId, latexBase64
    VTSetWordMetadataPayload documentObject, formulaId, encodedMetadata
    VTSetWordFormulaFormat documentObject, formulaId, displayMode, numbered
    VTSetWordImageScaleState _
        documentObject, formulaId, 14#, widthPoints, heightPoints, 0#, 14#
    Set formulaShape = VTEnsureVisualTeXImageMacroButton(formulaShape)

    If displayMode = "block" Then
        If numbered Then
            Set numberRange = VTInsertEquationNumber( _
                formulaShape, formulaId, _
                VTEquationCrossReferenceText(latexBase64))
            Set formulaShape = VTResolveImageFormulaInParagraph( _
                documentObject, numberRange.Paragraphs(1).Range.Start)
        Else
            VTNormalizeUnnumberedDisplayParagraph formulaShape.Range
            VTNormalizeImageDisplayParagraph formulaShape.Range
        End If
    End If
    VTSetWordImageOwnerBookmark documentObject, formulaShape, formulaId
    Set VTRegressionCreateGlobalFormatImage = formulaShape
End Function

Private Function VTRegressionGlobalFormatFormulaKind( _
    ByVal documentObject As Document, _
    ByVal formulaId As String) As String

    Dim formulaShape As InlineShape
    Dim nativeMath As OMath
    Dim nativeBookmarkName As String

    Set formulaShape = VTWordImageOwnerShape(documentObject, formulaId)
    If Not formulaShape Is Nothing Then
        VTRegressionGlobalFormatFormulaKind = "image"
        Exit Function
    End If
    nativeBookmarkName = VTNativeFormulaBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(nativeBookmarkName) Then
        Set nativeMath = VTNativeMathForBookmark( _
            documentObject.Bookmarks(nativeBookmarkName))
        If Not nativeMath Is Nothing Then
            VTRegressionGlobalFormatFormulaKind = "omml"
            Exit Function
        End If
    End If
    VTRegressionGlobalFormatFormulaKind = "missing"
End Function

Public Sub VisualTeX_SetupGlobalFormatConversionRegression()
    Const formulaId1 As String = _
        "61616161-6161-4616-8616-616161616161"
    Const formulaId2 As String = _
        "62626262-6262-4626-8626-626262626262"
    Const formulaId3 As String = _
        "63636363-6363-4636-8636-636363636363"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim fixtureRoot As String
    Dim metadata1 As String
    Dim metadata2 As String
    Dim metadata3 As String

    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    metadata1 = VTReadText( _
        fixtureRoot & "/word-global-format-metadata-1.txt", 131072)
    metadata2 = VTReadText( _
        fixtureRoot & "/word-global-format-metadata-2.txt", 131072)
    metadata3 = VTReadText( _
        fixtureRoot & "/word-global-format-metadata-3.txt", 131072)
    If Not VTIsEncodedMetadata(metadata1) Or _
       Not VTIsEncodedMetadata(metadata2) Or _
       Not VTIsEncodedMetadata(metadata3) Then
        Err.Raise vbObjectError + 7600, "VisualTeX regression", _
            "The global format conversion metadata fixtures are missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.Activate
    testDocument.ActiveWindow.View.Type = wdPrintView
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_MODE_VARIABLE, _
        VT_WORD_NUMBERING_MODE_SEQUENCE
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_SEPARATOR_VARIABLE, "."

    Set formulaShape = VTRegressionCreateGlobalFormatImage( _
        testDocument, formulaId1, metadata1, "x^2+1", _
        "inline", False, 60!, 20!)
    Set formulaShape = VTRegressionCreateGlobalFormatImage( _
        testDocument, formulaId2, metadata2, "\\frac{a}{b}", _
        "block", False, 76!, 28!)
    Set formulaShape = VTRegressionCreateGlobalFormatImage( _
        testDocument, formulaId3, metadata3, _
        "\\sum_{i=1}^{n}\\left(\\binom{n}{i}\\right)", _
        "block", True, 150!, 42!)
    formulaShape.Select
End Sub

Public Sub VisualTeX_InspectGlobalFormatConversionRegression()
    Const formulaId1 As String = _
        "61616161-6161-4616-8616-616161616161"
    Const formulaId2 As String = _
        "62626262-6262-4626-8626-626262626262"
    Const formulaId3 As String = _
        "63636363-6363-4636-8636-636363636363"

    Dim documentObject As Document
    Dim resultPath As String

    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-global-format-conversion-result.txt"
    If Documents.Count = 0 Then
        VTWriteTextAtomic resultPath, _
            "FAIL" & vbLf & "reason=no-document" & vbLf
        Exit Sub
    End If
    Set documentObject = ActiveDocument
    VTWriteTextAtomic resultPath, _
        "kind1=" & VTRegressionGlobalFormatFormulaKind( _
            documentObject, formulaId1) & vbLf & _
        "kind2=" & VTRegressionGlobalFormatFormulaKind( _
            documentObject, formulaId2) & vbLf & _
        "kind3=" & VTRegressionGlobalFormatFormulaKind( _
            documentObject, formulaId3) & vbLf & _
        "images=" & CStr(documentObject.InlineShapes.Count) & vbLf & _
        "omml=" & CStr(documentObject.OMaths.Count) & vbLf & _
        "tables=" & CStr(documentObject.Tables.Count) & vbLf & _
        "sequences=" & CStr(VTCountManagedEquationSequences( _
            documentObject)) & vbLf & _
        "number3=" & VTEquationNumberTextForFormula( _
            documentObject, formulaId3) & vbLf
End Sub

Public Sub VisualTeX_SetupComplexNumberingRibbonRegression()
    Const formulaId As String = _
        "99999999-9999-4999-8999-999999999999"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim ommlBase64 As String

    ommlBase64 = VTReadText( _
        VTApplicationSupportRoot() & _
            "/Tests/word-native-regression-complex-numbering-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    If Len(ommlBase64) = 0 Or _
       Not VTPathFileExists(VTNativeWordDocumentPath(formulaId)) Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The complex numbered Ribbon fixture is missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.Activate
    testDocument.ActiveWindow.View.Type = wdPrintView
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_MODE_VARIABLE, _
        VT_WORD_NUMBERING_MODE_SEQUENCE
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_SEPARATOR_VARIABLE, "."
    Set formulaShape = VTRegressionCreateNumberedImage( _
        testDocument, formulaId, _
        VTBase64UrlEncodeUtf8( _
            "\\sum_{i=1}^{n}\\left(\\binom{n}{i}\\right)"), _
        150!, 42!)
    VTSetWordOmmlPayload testDocument, formulaId, ommlBase64
    VTSetWordImageScaleState _
        testDocument, formulaId, 14#, 150#, 42#, 0#, 14#
    formulaShape.Select
End Sub

Public Sub VisualTeX_InspectComplexNumberingRibbonRegression()
    Const formulaId As String = _
        "99999999-9999-4999-8999-999999999999"

    Dim documentObject As Document
    Dim formulaRange As Range
    Dim nativeMath As OMath
    Dim visibleNumberField As Field
    Dim resultPath As String
    Dim hasArrayMarker As Boolean
    Dim hasVisibleRef As Boolean

    If Documents.Count = 0 Then Exit Sub
    Set documentObject = ActiveDocument
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-complex-numbering-ribbon-result.txt"
    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If formulaRange Is Nothing Or formulaRange.OMaths.Count <> 1 Then
        VTWriteTextAtomic resultPath, "FAIL" & vbLf & "reason=missing-omml" & vbLf
        Exit Sub
    End If
    Set nativeMath = formulaRange.OMaths(1)
    hasArrayMarker = Not _
        VTNativeEquationArrayMarkerRange(nativeMath.Range.Duplicate) Is Nothing
    Set visibleNumberField = VTNativeEquationArrayReferenceField( _
        nativeMath.Range.Duplicate, formulaId)
    If Not visibleNumberField Is Nothing Then
        hasVisibleRef = VTNativeEquationNumberIsInsideMath( _
            nativeMath.Range.Duplicate, visibleNumberField)
    End If
    VTWriteTextAtomic resultPath, _
        "kind=omml" & vbLf & _
        "number=" & VTEquationNumberTextForFormula( _
            documentObject, formulaId) & vbLf & _
        "tables=" & CStr(documentObject.Tables.Count) & vbLf & _
        "sequences=" & CStr(VTCountManagedEquationSequences( _
            documentObject)) & vbLf & _
        "visibleRef=" & CStr(hasVisibleRef) & vbLf & _
        "arrayMarker=" & CStr(hasArrayMarker) & vbLf
End Sub

Public Sub VisualTeX_RunImageToNativeFastPerformanceRegression()
    Const formulaId As String = _
        "11111111-1111-4111-8111-111111111111"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim nativeMath As OMath
    Dim ommlBase64 As String
    Dim conversionPath As String
    Dim resultPath As String
    Dim phaseStartedAt As Single
    Dim elapsedSeconds As Double
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-image-to-native-fast-performance-result.txt"
    On Error GoTo RegressionFailed

    If Not VTPathFileExists(VTNativeWordDocumentPath(formulaId)) Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The image-to-native performance cache fixture is missing."
    End If
    ommlBase64 = VTReadText( _
        VTApplicationSupportRoot() & _
            "/Tests/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    If Len(ommlBase64) = 0 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The image-to-native performance OMML fixture is missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.Activate
    testDocument.ActiveWindow.View.Type = wdPrintView
    Set formulaShape = VTRegressionCreateGlobalFormatImage( _
        testDocument, formulaId, VT_METADATA_PREFIX & "e30", "x", _
        "block", False, 63.01575!, 33.25063!)
    VTSetWordOmmlPayload testDocument, formulaId, ommlBase64

    phaseStartedAt = Timer
    If Not VTWordConvertInlineShapeToNativeFast( _
       formulaShape, conversionPath) Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The single image did not use the cached image-to-OMML path."
    End If
    elapsedSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)

    If conversionPath <> "direct-cache-insert" Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The single display image reopened a hidden staging document" & _
            " [path=" & conversionPath & "]."
    End If
    If testDocument.InlineShapes.Count <> 0 Or _
       testDocument.OMaths.Count <> 1 Or _
       testDocument.Tables.Count <> 0 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The direct image-to-OMML path did not leave one native formula."
    End If
    If Not testDocument.Bookmarks.Exists( _
       VTNativeFormulaBookmarkName(formulaId)) Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The direct image-to-OMML path lost the formula identity."
    End If
    Set nativeMath = VTNativeMathForBookmark(testDocument.Bookmarks( _
        VTNativeFormulaBookmarkName(formulaId)))
    If nativeMath Is Nothing Or nativeMath.Type <> wdOMathDisplay Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The direct image-to-OMML path did not preserve display mode."
    End If

    VTWriteTextAtomic resultPath, _
        "PASS" & vbLf & _
        "conversionPath=" & conversionPath & vbLf & _
        "imageToNativeMs=" & _
            CStr(CLng(elapsedSeconds * 1000#)) & vbLf & _
        "images=" & CStr(testDocument.InlineShapes.Count) & vbLf & _
        "omml=" & CStr(testDocument.OMaths.Count) & vbLf
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic resultPath, _
        "FAIL" & vbLf & _
        "conversionPath=" & conversionPath & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            regressionErrorDescription, vbCr, " "), vbLf, " ") & vbLf
    If Not testDocument Is Nothing Then _
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX image-to-native fast performance regression", _
        regressionErrorDescription
End Sub

Public Sub VisualTeX_RunNumberedImageOmmlFastPerformanceRegression()
    Const formulaId As String = _
        "0274750e-c676-4f7b-8c04-a9d9ef7336a0"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim insertionRange As Range
    Dim nativeBookmark As Bookmark
    Dim nativeMath As OMath
    Dim placeRef As Field
    Dim sequenceField As Field
    Dim visibleNumberField As Field
    Dim ommlBase64 As String
    Dim numberText As String
    Dim encodedMetadata As String
    Dim resultPath As String
    Dim resultText As String
    Dim imageToNativePath As String
    Dim imageToNativePhases As Variant
    Dim nativeToImagePhases As Variant
    Dim phaseStartedAt As Single
    Dim imageToNativeSeconds As Double
    Dim nativeToImageSeconds As Double
    Dim sourceWidthPoints As Double
    Dim sourceHeightPoints As Double
    Dim finalWidthPoints As Double
    Dim finalHeightPoints As Double
    Dim sourceAspectRatio As Double
    Dim finalAspectRatio As Double
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-numbered-image-omml-fast-performance-result.txt"
    On Error GoTo RegressionFailed

    If Not VTPathFileExists(VTNativeWordDocumentPath(formulaId)) Or _
       Not VTPathFileExists(VTImageCacheDocumentPath(formulaId)) Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The numbered conversion performance cache fixture is missing."
    End If
    ommlBase64 = VTReadText( _
        VTApplicationSupportRoot() & _
            "/Tests/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    If Len(ommlBase64) = 0 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The numbered conversion OMML fixture is missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.Activate
    testDocument.ActiveWindow.View.Type = wdPrintView
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_MODE_VARIABLE, _
        VT_WORD_NUMBERING_MODE_SEQUENCE
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_SEPARATOR_VARIABLE, "."
    encodedMetadata = VT_METADATA_PREFIX & "e30"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set formulaShape = VTAddWordFormulaPicture( _
        testDocument, insertionRange, _
        VTImageCacheDocumentPath(formulaId), _
        VTImageCachePngPath(formulaId))
    sourceWidthPoints = formulaShape.Width
    sourceHeightPoints = formulaShape.Height
    If sourceWidthPoints <= 0# Or sourceHeightPoints <= 0# Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The numbered conversion image cache has invalid geometry."
    End If
    sourceAspectRatio = sourceWidthPoints / sourceHeightPoints
    formulaShape.AlternativeText = encodedMetadata
    formulaShape.Title = VTFormulaReference(formulaId, "block", True)
    VTSetWordLatexPayload _
        testDocument, formulaId, _
        VTBase64UrlEncodeUtf8( _
            "x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}")
    VTSetWordMetadataPayload _
        testDocument, formulaId, encodedMetadata
    VTSetWordFormulaFormat testDocument, formulaId, "block", True
    Set formulaRange = VTInsertEquationNumber( _
        formulaShape, formulaId, _
        "x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}")
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    If formulaRange Is Nothing Or _
       formulaRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The numbered conversion fixture lost its image formula."
    End If
    Set formulaShape = formulaRange.InlineShapes(1)
    VTSetWordOmmlPayload testDocument, formulaId, ommlBase64
    VTSetWordImageScaleState _
        testDocument, formulaId, 14#, sourceWidthPoints, _
        sourceHeightPoints, 0#, 14#
    Set placeRef = VTImagePlaceRef(formulaShape.Range)
    Set sequenceField = VTImageSequenceField(placeRef, False)
    If placeRef Is Nothing Or sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The numbered image fixture lost its private native sequence."
    End If
    numberText = VTEquationNumberTextForFormula(testDocument, formulaId)
    If Len(numberText) = 0 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The numbered image fixture has no visible number."
    End If

    phaseStartedAt = Timer
    If Not VTWordConvertInlineShapeToNativeFast( _
       formulaShape, imageToNativePath, imageToNativePhases) Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The numbered image did not use the cached image-to-OMML fast path."
    End If
    imageToNativeSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    If imageToNativePath <> "numbered-direct-cache-insert" Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The numbered image did not use direct cached DOCX insertion" & _
            " [path=" & imageToNativePath & "]."
    End If

    If testDocument.InlineShapes.Count <> 0 Or _
       testDocument.OMaths.Count <> 1 Or _
       testDocument.Tables.Count <> 0 Or _
       VTCountManagedEquationSequences(testDocument) <> 1 Or _
        VTEquationNumberTextForFormula(testDocument, formulaId) <> numberText Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Numbered image-to-OMML fast conversion changed the canonical numbering structure" & _
            " [images=" & CStr(testDocument.InlineShapes.Count) & _
            "; omml=" & CStr(testDocument.OMaths.Count) & _
            "; tables=" & CStr(testDocument.Tables.Count) & _
            "; sequences=" & CStr(VTCountManagedEquationSequences(testDocument)) & _
            "; before=" & numberText & _
            "; after=" & VTEquationNumberTextForFormula( _
                testDocument, formulaId) & "]."
    End If
    Set nativeBookmark = testDocument.Bookmarks( _
        VTNativeFormulaBookmarkName(formulaId))
    Set nativeMath = VTNativeMathForBookmark(nativeBookmark)
    If nativeMath Is Nothing Or _
       VTNativeEquationArrayMarkerRange(nativeMath.Range.Duplicate) Is Nothing Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Numbered image-to-OMML fast conversion did not create the canonical Equation array."
    End If
    Set visibleNumberField = VTPureNativeNumberField( _
        nativeMath.Range.Duplicate)
    If visibleNumberField Is Nothing Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The numbered OMML fast path lost its internal native SEQ."
    End If

    nativeBookmark.Range.Select
    phaseStartedAt = Timer
    If Not VTWordConvertNativeBookmarkToImageFast( _
       nativeBookmark, nativeToImagePhases) Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The numbered OMML did not use the cached native-to-image fast path."
    End If
    nativeToImageSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)

    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    If formulaRange Is Nothing Or _
       formulaRange.InlineShapes.Count <> 1 Or _
       formulaRange.OMaths.Count <> 0 Or _
       formulaRange.Information(wdWithInTable) Or _
       testDocument.Tables.Count <> 0 Or _
       VTCountManagedEquationSequences(testDocument) <> 1 Or _
       VTEquationNumberTextForFormula(testDocument, formulaId) <> numberText Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "Numbered OMML-to-image fast conversion changed the canonical numbering structure."
    End If
    Set placeRef = VTImagePlaceRef(formulaRange)
    Set visibleNumberField = VTImageSequenceField(placeRef, False)
    If placeRef Is Nothing Or visibleNumberField Is Nothing Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The numbered image fast roundtrip lost its private sequence."
    End If
    Set formulaShape = formulaRange.InlineShapes(1)
    finalWidthPoints = formulaShape.Width
    finalHeightPoints = formulaShape.Height
    If finalWidthPoints <= 0# Or finalHeightPoints <= 0# Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The numbered conversion roundtrip produced invalid image geometry."
    End If
    finalAspectRatio = finalWidthPoints / finalHeightPoints
    If Abs(finalWidthPoints - sourceWidthPoints) > 0.1 Or _
       Abs(finalHeightPoints - sourceHeightPoints) > 0.1 Or _
       Abs(finalAspectRatio - sourceAspectRatio) > 0.01 Then
        Err.Raise vbObjectError + 7599, "VisualTeX regression", _
            "The numbered conversion roundtrip compressed the formula image" & _
            " [source=" & CStr(sourceWidthPoints) & "x" & _
                CStr(sourceHeightPoints) & _
            "; final=" & CStr(finalWidthPoints) & "x" & _
                CStr(finalHeightPoints) & "]."
    End If

    resultText = "PASS" & vbLf
    resultText = resultText & "imageToNativeMs=" & _
        CStr(CLng(imageToNativeSeconds * 1000#)) & vbLf
    resultText = resultText & "nativeToImageMs=" & _
        CStr(CLng(nativeToImageSeconds * 1000#)) & vbLf
    resultText = resultText & _
        "imageToNativePath=" & imageToNativePath & vbLf
    resultText = resultText & "imagePrepareMs=" & _
        CStr(CLng(CDbl(imageToNativePhases(0)) * 1000#)) & vbLf
    resultText = resultText & "imageInsertMs=" & _
        CStr(CLng(CDbl(imageToNativePhases(1)) * 1000#)) & vbLf
    resultText = resultText & "imageDeleteMs=" & _
        CStr(CLng(CDbl(imageToNativePhases(2)) * 1000#)) & vbLf
    resultText = resultText & "imageNumberMs=" & _
        CStr(CLng(CDbl(imageToNativePhases(3)) * 1000#)) & vbLf
    resultText = resultText & "imageFinalizeMs=" & _
        CStr(CLng(CDbl(imageToNativePhases(4)) * 1000#)) & vbLf
    resultText = resultText & "nativeResolveMs=" & _
        CStr(CLng(CDbl(nativeToImagePhases(0)) * 1000#)) & vbLf
    resultText = resultText & "nativeMetadataMs=" & _
        CStr(CLng(CDbl(nativeToImagePhases(1)) * 1000#)) & vbLf
    resultText = resultText & "nativeCacheMs=" & _
        CStr(CLng(CDbl(nativeToImagePhases(2)) * 1000#)) & vbLf
    resultText = resultText & "nativeTransactionMs=" & _
        CStr(CLng(CDbl(nativeToImagePhases(3)) * 1000#)) & vbLf
    resultText = resultText & "nativeStagingMs=" & _
        CStr(CLng(CDbl(nativeToImagePhases(4)) * 1000#)) & vbLf
    resultText = resultText & "nativeFinalizeMs=" & _
        CStr(CLng(CDbl(nativeToImagePhases(5)) * 1000#)) & vbLf
    resultText = resultText & "number=" & numberText & vbLf
    resultText = resultText & _
        "tables=" & CStr(testDocument.Tables.Count) & vbLf
    resultText = resultText & "sequences=" & _
        CStr(VTCountManagedEquationSequences(testDocument)) & vbLf
    resultText = resultText & "sourceImageSize=" & _
        CStr(sourceWidthPoints) & "x" & CStr(sourceHeightPoints) & vbLf
    resultText = resultText & "finalImageSize=" & _
        CStr(finalWidthPoints) & "x" & CStr(finalHeightPoints) & vbLf
    VTWriteTextAtomic resultPath, resultText
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic resultPath, _
        "FAIL" & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
    If Not testDocument Is Nothing Then _
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    On Error GoTo 0
    Exit Sub
End Sub

Public Sub VisualTeX_RunWordProductionCarrierRegression()
    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim nativeEquation As OMath
    Dim dispatch As Collection
    Dim sequenceField As Field
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim latexBase64 As String
    Dim ommlBase64 As String
    Dim metadata As String
    Dim fixtureRoot As String
    Dim sourcePath As String
    Dim sessionId As String
    Dim sourceMarker As String
    Dim stage As String
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim observedFontSizePt As Double
    Dim numberBefore As String

    On Error GoTo Failed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    sourcePath = VTReadText(fixtureRoot & "/word-production-carrier-input.txt", 4096)
    stage = "open-production-snapshot-copy"
    Set testDocument = Documents.Open(FileName:=Trim$(sourcePath), _
        AddToRecentFiles:=False, Visible:=True)
    testDocument.Activate
    If testDocument.InlineShapes.Count <> 2 Then Err.Raise 5, , "Expected two source images."
    Set formulaShape = testDocument.InlineShapes(2)
    If Not VTTryParseFormulaReference(formulaShape.Title, formulaId, displayMode, numbered) Then Err.Raise 5
    metadata = formulaShape.AlternativeText
    If Not VTTryReadWordLatexPayload(testDocument, formulaId, latexBase64) Then Err.Raise 5
    If Not VTTryReadWordOmmlPayload(testDocument, formulaId, ommlBase64) Then Err.Raise 5
    VTReadWordImageConversionState formulaShape, formulaId, fontSizePt, _
        referenceWidthPt, referenceHeightPt, referenceBaselinePt, observedFontSizePt
    Set sequenceField = VTEquationSequenceFieldForBookmark(testDocument, _
        VTEquationSequenceNumberBookmarkName(formulaId))
    numberBefore = sequenceField.Result.Text

    ' Enter the same callback as a renderer response after native-cache rebuild.
    ' The edit bookmark deliberately points to the production image-in-OMath.
    stage = "resume-production-image-carrier"
    sessionId = VTNewUuidV4()
    sourceMarker = VTWordEditBookmarkName(sessionId)
    VTAddWordEditRangeBookmark formulaShape.Range, sessionId
    Set dispatch = New Collection
    dispatch.Add "edit", "mode"
    dispatch.Add formulaId, "formulaId"
    dispatch.Add displayMode, "displayMode"
    dispatch.Add "1", "numbered"
    dispatch.Add "1", "nativeEquation"
    dispatch.Add "imageToNative", "operation"
    dispatch.Add metadata, "metadata"
    dispatch.Add latexBase64, "latexBase64"
    dispatch.Add ommlBase64, "ommlBase64"
    dispatch.Add VTNativeWordDocumentPath(formulaId), "nativeDocumentPath"
    dispatch.Add sourceMarker, "sourceMarker"
    dispatch.Add VTWordDocumentIdentityForDocument(testDocument), "sourceDocumentId"
    dispatch.Add CStr(formulaShape.Width), "widthPoints"
    dispatch.Add CStr(formulaShape.Height), "heightPoints"
    dispatch.Add CStr(fontSizePt), "fontSizePt"
    dispatch.Add CStr(referenceWidthPt), "referenceWidthPt"
    dispatch.Add CStr(referenceHeightPt), "referenceHeightPt"
    dispatch.Add CStr(referenceBaselinePt), "referenceBaselinePt"
    VTCommitWordDispatch sessionId, dispatch
    If testDocument.OMaths.Count <> 1 Or testDocument.InlineShapes.Count <> 1 Then Err.Raise 5, , "Resume duplicated or lost the carrier."
    If testDocument.Tables.Count <> 0 Then Err.Raise 5, , "Table fallback."
    Set nativeEquation = VTNativeMathForBookmark(testDocument.Bookmarks(VTNativeFormulaBookmarkName(formulaId)))
    If nativeEquation Is Nothing Then Err.Raise 5, , "Missing native identity."
    Set sequenceField = VTNativeEquationSequenceHelperField(testDocument, formulaId)
    If sequenceField.Result.Text <> numberBefore Then Err.Raise 5, , "Sequence changed."
    VTWriteTextAtomic fixtureRoot & "/word-production-carrier-native.xml", testDocument.Content.WordOpenXML

    stage = "production-native-to-image"
    If Not VTWordConvertNativeBookmarkToImageFast(testDocument.Bookmarks(VTNativeFormulaBookmarkName(formulaId))) Then Err.Raise 5, , "Image cache unavailable."
    If testDocument.OMaths.Count <> 0 Or testDocument.InlineShapes.Count <> 2 Then Err.Raise 5, , "Image retains a math zone."
    Set formulaShape = testDocument.InlineShapes(2)
    If formulaShape.Range.Paragraphs(1).Range.OMaths.Count <> 0 Then Err.Raise 5
    VTWriteTextAtomic fixtureRoot & "/word-production-carrier-image.xml", testDocument.Content.WordOpenXML

    stage = "ordinary-numbered-image-to-native"
    If Not VTWordConvertInlineShapeToNativeFast(formulaShape) Then Err.Raise 5, , "Canonical fast path unavailable."
    If testDocument.OMaths.Count <> 1 Or testDocument.InlineShapes.Count <> 1 Then Err.Raise 5
    If testDocument.Tables.Count <> 0 Then Err.Raise 5
    Set sequenceField = VTNativeEquationSequenceHelperField(testDocument, formulaId)
    If sequenceField.Result.Text <> numberBefore Then Err.Raise 5, , "Round trip changed number."
    VTWriteTextAtomic fixtureRoot & "/word-production-carrier-roundtrip.xml", testDocument.Content.WordOpenXML
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    VTWriteTextAtomic fixtureRoot & "/word-production-carrier-result.txt", "PASS" & vbLf & _
        "resumeCallback=PASS" & vbLf & "nativeToPlainImage=PASS" & vbLf & "ordinaryImageToNative=PASS" & vbLf
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Not testDocument Is Nothing Then
        VTWriteTextAtomic fixtureRoot & "/word-production-carrier-failure.xml", testDocument.Content.WordOpenXML
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    VTWriteTextAtomic fixtureRoot & "/word-production-carrier-result.txt", "FAIL" & vbLf & _
        "stage=" & stage & vbLf & "error=" & CStr(failureNumber) & ": " & failureDescription
End Sub

Public Sub VisualTeX_StartWordProductionColdCacheRegression()
    Const formulaId As String = _
        "cb753a13-bf06-447e-a0bf-41cbae730cf5"
    Const expectedNumberVariable As String = _
        "VTRegressionColdCacheNumber"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim candidateShape As InlineShape
    Dim sequenceField As Field
    Dim candidateId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim fixtureRoot As String
    Dim sourcePath As String
    Dim resultPath As String
    Dim numberBefore As String
    Dim failureNumber As Long
    Dim failureDescription As String

    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    resultPath = fixtureRoot & _
        "/word-production-cold-cache-result.txt"
    On Error GoTo Failed
    If VTPathFileExists(VTNativeWordDocumentPath(formulaId)) Then
        Err.Raise vbObjectError + 7601, "VisualTeX regression", _
            "The production cold-cache fixture unexpectedly has a native cache."
    End If
    sourcePath = VTReadText( _
        fixtureRoot & "/word-production-carrier-input.txt", 4096)
    Set testDocument = Documents.Open(FileName:=Trim$(sourcePath), _
        AddToRecentFiles:=False, Visible:=True)
    testDocument.AttachedTemplate = ThisDocument.FullName
    testDocument.UpdateStylesOnOpen = False
    testDocument.Activate
    For Each candidateShape In testDocument.InlineShapes
        candidateId = ""
        If VTTryParseFormulaReference( _
           candidateShape.Title, candidateId, displayMode, numbered) Then
            If candidateId = formulaId Then
                Set formulaShape = candidateShape
                Exit For
            End If
        End If
    Next candidateShape
    If formulaShape Is Nothing Or _
       formulaShape.Range.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7601, "VisualTeX regression", _
            "The production cold-cache source is not one ordinary image."
    End If
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, formulaId)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7601, "VisualTeX regression", _
            "The production cold-cache source has no external Equation helper."
    End If
    numberBefore = sequenceField.Result.Text
    VTSetDocumentVariable _
        testDocument, expectedNumberVariable, numberBefore
    VTWriteTextAtomic resultPath, _
        "STARTED" & vbLf & "number=" & numberBefore & vbLf
    VTWordConvertInlineShapeToNativeEquation formulaShape
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic resultPath, _
        "FAIL" & vbLf & _
        "stage=start" & vbLf & _
        "error=" & CStr(failureNumber) & ": " & failureDescription & vbLf
    If Not testDocument Is Nothing Then _
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    On Error GoTo 0
End Sub

Public Sub VisualTeX_InspectWordProductionColdCacheRegression()
    Const formulaId As String = _
        "cb753a13-bf06-447e-a0bf-41cbae730cf5"
    Const expectedNumberVariable As String = _
        "VTRegressionColdCacheNumber"

    Dim testDocument As Document
    Dim nativeEquation As OMath
    Dim sequenceField As Field
    Dim markerRange As Range
    Dim fixtureRoot As String
    Dim resultPath As String
    Dim expectedNumber As String
    Dim stage As String
    Dim failureNumber As Long
    Dim failureDescription As String

    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    resultPath = fixtureRoot & _
        "/word-production-cold-cache-result.txt"
    On Error GoTo Failed
    stage = "resolve-document"
    Set testDocument = ActiveDocument
    If Not VTTryGetDocumentVariable( _
       testDocument, expectedNumberVariable, expectedNumber) Then
        Err.Raise vbObjectError + 7601, "VisualTeX regression", _
            "The active document is not the production cold-cache fixture."
    End If
    stage = "validate-native-cache"
    If Not VTPathFileExists(VTNativeWordDocumentPath(formulaId)) Then
        Err.Raise vbObjectError + 7601, "VisualTeX regression", _
            "The renderer did not persist the rebuilt native cache."
    End If
    stage = "validate-native-bookmark"
    If Not testDocument.Bookmarks.Exists( _
       VTNativeFormulaBookmarkName(formulaId)) Then
        Err.Raise vbObjectError + 7601, "VisualTeX regression", _
            "The cold-cache callback did not create a native formula identity."
    End If
    Set nativeEquation = VTNativeMathForBookmark( _
        testDocument.Bookmarks(VTNativeFormulaBookmarkName(formulaId)))
    If nativeEquation Is Nothing Or _
       nativeEquation.Range.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 7601, "VisualTeX regression", _
            "The cold-cache callback did not create ordinary numbered OMML."
    End If
    Set markerRange = VTNativeEquationArrayMarkerRange( _
        nativeEquation.Range.Duplicate)
    If markerRange Is Nothing Or _
       InStr(1, nativeEquation.Range.Text, expectedNumber, _
           vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7601, "VisualTeX regression", _
            "The cold-cache callback lost the canonical Equation-array number."
    End If
    stage = "validate-external-sequence"
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, formulaId)
    If sequenceField Is Nothing Or _
       sequenceField.Result.Text <> expectedNumber Then
        Err.Raise vbObjectError + 7601, "VisualTeX regression", _
            "The cold-cache callback changed the Equation ordinal."
    End If
    If testDocument.Tables.Count <> 0 Or _
       testDocument.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7601, "VisualTeX regression", _
            "The cold-cache callback duplicated a carrier or used a table fallback."
    End If
    stage = "write-result"
    VTWriteTextAtomic resultPath, _
        "PASS" & vbLf & _
        "rendererCacheRebuild=PASS" & vbLf & _
        "callbackResume=PASS" & vbLf & _
        "number=" & expectedNumber & vbLf
    stage = "close-document"
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic resultPath, _
        "FAIL" & vbLf & _
        "stage=" & stage & vbLf & _
        "error=" & CStr(failureNumber) & ": " & failureDescription & vbLf
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
End Sub

Public Sub VisualTeX_RunWordImageNativeFallbackRegression()
    ' Build the regression dynamically so the acceptance target is the current
    ' canonical numbered-image layout rather than a historical fallback DOCX.
    ' The second formula uses professional n-ary + delimiter + noBar/binomial
    ' OMML, the combination that used to trigger runtime error 445 on a redundant
    ' source BuildUp before Equation-array numbering.
    Dim testDocument As Document
    Dim sourceImage As InlineShape
    Dim nativeEquation As OMath
    Dim sequenceField As Field
    Dim markerRange As Range
    Dim dispatch As Collection
    Dim fixtureRoot As String
    Dim formulaId As String
    Dim firstFormulaId As String
    Dim complexLatexBase64 As String
    Dim complexOmmlBase64 As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim payload As String
    Dim sourceMetadata As String
    Dim beforeText As String
    Dim numberText As String
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim observedFontSizePt As Double
    Dim sourceStart As Long
    Dim sourceHeight As Double
    Dim testIndex As Long
    Dim expectedError As Long
    Dim expectedDescription As String
    Dim expectedSource As String
    Dim stage As String
    Dim failureNumber As Long
    Dim failureDescription As String

    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    formulaId = "99999999-9999-4999-8999-999999999999"
    firstFormulaId = "98989898-9898-4989-8989-989898989898"
    complexLatexBase64 = VTBase64UrlEncodeUtf8( _
        "\\sum_{i=1}^{n}\\left(\\binom{n}{i}\\right)")
    complexOmmlBase64 = VTReadText( _
        fixtureRoot & "/word-native-regression-complex-numbering-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    If Len(complexOmmlBase64) = 0 Or _
       Not VTPathFileExists(VTNativeWordDocumentPath(formulaId)) Then
        Err.Raise vbObjectError + 7564, "VisualTeX regression", _
            "The complex numbered OMML fixture is missing."
    End If

    On Error GoTo Failed
    For testIndex = 1 To 2
        stage = "create-fixture-" & CStr(testIndex)
        Set testDocument = Documents.Add(Visible:=True)
        testDocument.Activate
        testDocument.ActiveWindow.View.Type = wdPrintView
        VTSetDocumentVariable _
            testDocument, VT_WORD_NUMBERING_MODE_VARIABLE, _
            VT_WORD_NUMBERING_MODE_SEQUENCE
        VTSetDocumentVariable _
            testDocument, VT_WORD_NUMBERING_SEPARATOR_VARIABLE, "."
        VTRegressionCreateNumberedImage _
            testDocument, firstFormulaId, VTBase64UrlEncodeUtf8("x"), _
            90!, 28!
        Set sourceImage = VTRegressionCreateNumberedImage( _
            testDocument, formulaId, complexLatexBase64, 150!, 42!)
        sourceMetadata = _
            VT_METADATA_PREFIX & VTBase64UrlEncodeUtf8(formulaId)
        sourceImage.AlternativeText = sourceMetadata
        VTSetWordMetadataPayload testDocument, formulaId, sourceMetadata
        VTSetWordImageOwnerBookmark testDocument, sourceImage, formulaId
        VTSetWordOmmlPayload testDocument, formulaId, complexOmmlBase64
        If testDocument.InlineShapes.Count <> 2 Or _
           testDocument.OMaths.Count <> 0 Then
            Err.Raise vbObjectError + 7564, "VisualTeX regression", _
                "The dynamic complex-numbering fixture is invalid."
        End If
        If Not VTTryParseFormulaReference( _
           sourceImage.Title, formulaId, displayMode, numbered) Then
            Err.Raise vbObjectError + 7564, "VisualTeX regression", _
                "Missing dynamic complex formula identity."
        End If
        sourceMetadata = sourceImage.AlternativeText
        Set sequenceField = VTNativeEquationSequenceHelperField(testDocument, formulaId)
        If sequenceField Is Nothing Then
            Err.Raise vbObjectError + 7564, "VisualTeX regression", _
                "The complex fixture lost its external SEQ identity."
        End If
        numberText = sequenceField.Result.Text
        ' Creating and formatting the external SEQ/REF structure can invalidate
        ' the InlineShape COM wrapper returned before those mutations. Model the
        ' Ribbon entry accurately by fetching the selected formula again from
        ' Word's current InlineShapes collection immediately before conversion.
        Set sourceImage = testDocument.InlineShapes(2)
        If sourceImage Is Nothing Then
            Err.Raise vbObjectError + 7564, "VisualTeX regression", _
                "The complex fixture lost its image owner before conversion."
        End If

        If testIndex = 1 Then
            stage = "source-deletion-rollback"
            beforeText = testDocument.Content.Text
            On Error Resume Next
            Call VTWordConvertInlineShapeToNativeEquation(sourceImage, True)
            expectedError = Err.Number
            expectedDescription = Err.Description
            expectedSource = Err.Source
            Err.Clear
            On Error GoTo Failed
            If expectedError = 0 Or InStr(expectedDescription, "Injected") = 0 Then
                Err.Raise vbObjectError + 7564, "VisualTeX regression", _
                    "The rollback probe did not reach source deletion" & _
                    " [source=" & expectedSource & "]: " & expectedDescription
            End If
            If testDocument.InlineShapes.Count <> 2 Or _
               testDocument.OMaths.Count <> 0 Then
                Err.Raise vbObjectError + 7564, "VisualTeX regression", _
                    "A layout failure did not restore the two image carriers" & _
                    " [images=" & CStr(testDocument.InlineShapes.Count) & _
                    "; omml=" & CStr(testDocument.OMaths.Count) & "]."
            End If
            If testDocument.Content.Text <> beforeText Or _
               testDocument.InlineShapes(2).AlternativeText <> sourceMetadata Or _
               VTCountManagedEquationSequences(testDocument) <> 2 Then
                Err.Raise vbObjectError + 7564, "VisualTeX regression", _
                    "A layout failure after source deletion did not restore the image and numbering."
            End If
            Set sequenceField = VTNativeEquationSequenceHelperField( _
                testDocument, formulaId)
            If sequenceField Is Nothing Or _
               sequenceField.Result.Text <> numberText Then
                Err.Raise vbObjectError + 7564, "VisualTeX regression", _
                    "The source-deletion rollback changed the Equation identity."
            End If
            VTWriteTextAtomic _
                fixtureRoot & "/word-image-native-complex-rollback.xml", _
                testDocument.Content.WordOpenXML
            Set sourceImage = testDocument.InlineShapes(2)
        End If

        stage = "complex-image-to-native-" & CStr(testIndex)
        VTWordConvertInlineShapeToNativeEquation sourceImage
        If testDocument.InlineShapes.Count <> 1 Or testDocument.OMaths.Count <> 1 Then
            Err.Raise vbObjectError + 7564, "VisualTeX regression", "Conversion lost or duplicated a formula."
        End If
        Set nativeEquation = VTNativeMathForBookmark( _
            testDocument.Bookmarks(VTNativeFormulaBookmarkName(formulaId)))
        If nativeEquation Is Nothing Then
            Err.Raise vbObjectError + 7564, "VisualTeX regression", "Conversion lost its native identity."
        End If
        If nativeEquation.Range.Information(wdWithInTable) Then
            Err.Raise vbObjectError + 7564, "VisualTeX regression", _
                "Complex native numbering regressed to the legacy table fallback."
        End If
        Set markerRange = VTNativeEquationArrayMarkerRange( _
            nativeEquation.Range.Duplicate)
        If markerRange Is Nothing Then
            Err.Raise vbObjectError + 7564, "VisualTeX regression", _
                "Complex native numbering did not create the canonical Equation-array marker."
        End If
        Set sequenceField = VTNativeEquationSequenceHelperField(testDocument, formulaId)
        If sequenceField Is Nothing Then
            Err.Raise vbObjectError + 7564, "VisualTeX regression", "Conversion lost its external helper."
        End If
        If sequenceField.Result.Text <> numberText Or _
           VTCountManagedEquationSequences(testDocument) <> 2 Then
            Err.Raise vbObjectError + 7564, "VisualTeX regression", "Conversion changed numbering."
        End If
        VTWriteTextAtomic _
            fixtureRoot & "/word-image-native-complex-" & _
                CStr(testIndex) & ".xml", _
            testDocument.Content.WordOpenXML
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set testDocument = Nothing
    Next testIndex
    VTWriteTextAtomic fixtureRoot & "/word-image-native-fallback-result.txt", _
        "PASS" & vbLf & _
        "sourceDeletionRollback=PASS" & vbLf & _
        "complexArrayNumbering=PASS" & vbLf & _
        "directComplexConversion=PASS" & vbLf
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Not testDocument Is Nothing Then
        VTWriteTextAtomic fixtureRoot & "/word-image-native-fallback-failure.xml", _
            testDocument.Content.WordOpenXML
    End If
    VTWriteTextAtomic fixtureRoot & "/word-image-native-fallback-result.txt", _
        "FAIL" & vbLf & "stage=" & stage & vbLf & _
        "errorNumber=" & CStr(failureNumber) & vbLf & _
        "errorDescription=" & failureDescription & vbLf
End Sub

Public Sub VisualTeX_RunWordNativeRegression()
    Const fixtureFormulaId As String = _
        "11111111-1111-4111-8111-111111111111"
    Const nativeFormulaId As String = _
        "22222222-2222-4222-8222-222222222222"
    Const conversionFormulaId As String = _
        "33333333-3333-4333-8333-333333333333"
    Const stabilityImageFormulaId As String = _
        "44444444-4444-4444-8444-444444444444"
    Const stabilityNativeFormulaId As String = _
        "55555555-5555-4555-8555-555555555555"

    Dim testDocument As Document
    Dim placeholder As InlineShape
    Dim referenceStabilityPlaceholder As InlineShape
    Dim insertionRange As Range
    Dim equationRange As Range
    Dim numberRange As Range
    Dim crossReferenceItems As Variant
    Dim nativeCrossReferenceItems As Variant
    Dim nativeDocumentPath As String
    Dim conversionDocumentPath As String
    Dim integralDocumentPath As String
    Dim sumFractionDocumentPath As String
    Dim ommlBase64 As String
    Dim fixtureRoot As String
    Dim equationStart As Long
    Dim itemCount As Long
    Dim itemIndex As Long
    Dim nativeEquationStart As Long
    Dim numberCreated As Boolean
    Dim sourceHeightPoints As Double
    Dim previousNumberText As String
    Dim diagnosticText As String
    Dim privateRefRepairDiagnostic As String
    Dim builtInTargetName As String
    Dim crossReferenceInventory As String
    Dim equationEndBefore As Long
    Dim equationEndAfter As Long
    Dim caretPositionBefore As Long
    Dim diagnosticIndex As Long
    Dim characterCode As Long
    Dim sequenceField As Field
    Dim referenceField As Field
    Dim builtInTargetRange As Range
    Dim newParagraph As Paragraph
    Dim diagnosticRange As Range
    Dim paragraphMarkRange As Range
    Dim crossReferenceTextFound As Boolean
    Dim crossReferenceNumberFound As Boolean
    Dim referenceResult As String
    Dim invariantSnapshot As String
    Dim invariantParagraphStart As Long
    Dim expectedUnnumberedPosition As Long
    Dim actualUnnumberedPosition As Long
    Dim paragraphMarkPosition As Long
    Dim inlineTableAdvance As Single
    Dim displayTableAdvance As Single
    Dim integralDisplayAdvance As Double
    Dim sumFractionDisplayAdvance As Double
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    nativeDocumentPath = _
        VTApplicationSupportRoot() & "/NativeDocuments/" & _
        fixtureFormulaId & ".docx"
    conversionDocumentPath = VTNativeWordDocumentPath(conversionFormulaId)
    integralDocumentPath = _
        VTApplicationSupportRoot() & _
        "/NativeDocuments/66666666-6666-4666-8666-666666666666.docx"
    sumFractionDocumentPath = _
        VTApplicationSupportRoot() & _
        "/NativeDocuments/77777777-7777-4777-8777-777777777777.docx"
    ommlBase64 = VTReadText( _
        fixtureRoot & "/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    If Not VTPathFileExists(nativeDocumentPath) Or _
       Not VTPathFileExists(conversionDocumentPath) Or _
       Not VTPathFileExists(integralDocumentPath) Or _
       Not VTPathFileExists(sumFractionDocumentPath) Then
        Err.Raise vbObjectError + 7482, "VisualTeX", _
            "A Word native regression DOCX fixture is missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_MODE_VARIABLE, _
        VT_WORD_NUMBERING_MODE_SEQUENCE
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_SEPARATOR_VARIABLE, "."

    regressionStage = "inline-after-source-removal"
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    Set placeholder = testDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Range:=insertionRange)
    placeholder.Width = 1
    placeholder.Height = 1
    placeholder.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Set insertionRange = placeholder.Range.Duplicate
    insertionRange.Collapse wdCollapseEnd
    Set equationRange = VTInsertNativeEquationAtRange( _
        insertionRange, ommlBase64, nativeDocumentPath, _
        "inline", False, False)
    equationStart = equationRange.Start
    placeholder.Delete
    Set equationRange = VTResolveNativeEquationRange( _
        testDocument, equationStart, 16)
    Set equationRange = VTFinalizeInlineNativeEquation(equationRange)
    If equationRange.OMaths.Count <> 1 Or _
       equationRange.OMaths(1).Type <> wdOMathInline Then
        Err.Raise vbObjectError + 7483, "VisualTeX", _
            "The empty-paragraph native equation did not remain inline."
    End If
    If equationRange.ParagraphFormat.Alignment <> wdAlignParagraphLeft Then
        Err.Raise vbObjectError + 7484, "VisualTeX", _
            "The empty-paragraph inline native equation remained centered."
    End If

    regressionStage = "inline-caret-adjacency"
    equationEndBefore = equationRange.End
    VTPlaceCaretAfterInlineNativeEquation equationRange
    caretPositionBefore = Selection.Start
    Selection.TypeText Text:="Z"
    Set equationRange = VTResolveNativeEquationRange( _
        testDocument, equationStart, 16)
    equationEndAfter = equationRange.End
    If equationEndAfter >= testDocument.Content.End Or _
       testDocument.Range( _
           Start:=equationEndAfter, End:=equationEndAfter + 1).Text <> "Z" Then
        diagnosticText = ""
        Set diagnosticRange = testDocument.Content.Duplicate
        For diagnosticIndex = 1 To Len(diagnosticRange.Text)
            characterCode = AscW(Mid$(diagnosticRange.Text, diagnosticIndex, 1))
            If characterCode < 0 Then characterCode = characterCode + 65536
            If Len(diagnosticText) > 0 Then diagnosticText = diagnosticText & ","
            diagnosticText = diagnosticText & CStr(characterCode)
        Next diagnosticIndex
        Err.Raise vbObjectError + 7503, "VisualTeX", _
            "Typing after an inline native equation introduced a separator. " & _
            "equationEndBefore=" & CStr(equationEndBefore) & _
            "; caretBefore=" & CStr(caretPositionBefore) & _
            "; equationEndAfter=" & CStr(equationEndAfter) & _
            "; selectionAfter=" & CStr(Selection.Start) & _
            "; characterCodes=" & diagnosticText
    End If
    If InStr(1, testDocument.Content.Text, ChrW(8288), _
       vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 7531, "VisualTeX", _
            "The empty-paragraph text anchor was not replaced by typed text."
    End If

    regressionStage = "inline-existing-reset"
    testDocument.Content.Delete
    regressionStage = "inline-existing-seed-text"
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    insertionRange.InsertAfter "A"
    Set insertionRange = testDocument.Range(Start:=1, End:=1)
    regressionStage = "inline-existing-placeholder"
    Set placeholder = testDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Range:=insertionRange)
    placeholder.Width = 1
    placeholder.Height = 1
    Set insertionRange = placeholder.Range.Duplicate
    insertionRange.Collapse wdCollapseEnd
    regressionStage = "inline-existing-insert-equation"
    Set equationRange = VTInsertNativeEquationAtRange( _
        insertionRange, ommlBase64, nativeDocumentPath, _
        "inline", False, False)
    equationStart = equationRange.Start
    regressionStage = "inline-existing-delete-source"
    placeholder.Delete
    regressionStage = "inline-existing-resolve"
    Set equationRange = VTResolveNativeEquationRange( _
        testDocument, equationStart, 16)
    regressionStage = "inline-existing-finalize"
    Set equationRange = VTFinalizeInlineNativeEquation(equationRange)
    regressionStage = "inline-existing-place-caret"
    VTPlaceCaretAfterInlineNativeEquation equationRange
    regressionStage = "inline-existing-type-text"
    Selection.TypeText Text:="Z"
    regressionStage = "inline-existing-resolve-after-type"
    Set equationRange = VTResolveNativeEquationRange( _
        testDocument, equationStart, 16)
    regressionStage = "inline-existing-assert"
    If equationRange.Start <= 0 Then
        Err.Raise vbObjectError + 7532, "VisualTeX", _
            "Inline OMML beside existing text lost its leading boundary."
    End If
    If equationRange.End >= testDocument.Content.End Then
        Err.Raise vbObjectError + 7532, "VisualTeX", _
            "Inline OMML beside existing text absorbed the following text."
    End If
    If testDocument.Range( _
           Start:=equationRange.Start - 1, _
           End:=equationRange.Start).Text <> "A" Or _
       testDocument.Range( _
           Start:=equationRange.End, _
           End:=equationRange.End + 1).Text <> "Z" Then
        Err.Raise vbObjectError + 7532, "VisualTeX", _
            "Inline OMML beside existing text changed its text boundaries."
    End If
    If InStr(1, testDocument.Content.Text, ChrW(8288), _
       vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 7533, "VisualTeX", _
            "Inline OMML beside existing text created an unnecessary anchor."
    End If
    inlineTableAdvance = VTOMathTableVisualAdvance( _
        equationRange, wdOMathInline, "inline-existing-table")

    regressionStage = "display-promotion-and-bookmark"
    testDocument.Content.Delete
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    Set placeholder = testDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Range:=insertionRange)
    placeholder.Width = 1
    placeholder.Height = 1
    Set insertionRange = placeholder.Range.Duplicate
    insertionRange.Collapse wdCollapseEnd
    Set equationRange = VTInsertNativeEquationAtRange( _
        insertionRange, ommlBase64, nativeDocumentPath, _
        "inline", True, False)
    equationStart = equationRange.Start
    placeholder.Delete
    Set equationRange = VTResolveNativeEquationRange( _
        testDocument, equationStart, 16)
    Set equationRange = VTPromoteNativeEquationToDisplay(equationRange)
    If equationRange.OMaths.Count <> 1 Or _
       equationRange.OMaths(1).Type <> wdOMathDisplay Then
        Err.Raise vbObjectError + 7485, "VisualTeX", _
            "The native display equation promotion failed."
    End If
    If equationRange.ParagraphFormat.Alignment <> wdAlignParagraphCenter Then
        Err.Raise vbObjectError + 7486, "VisualTeX", _
            "The native display equation was not centered."
    End If
    VTSetNativeFormulaBookmark _
        testDocument, equationRange, fixtureFormulaId
    If Not testDocument.Bookmarks.Exists( _
        VTNativeFormulaBookmarkName(fixtureFormulaId)) Then
        Err.Raise vbObjectError + 7487, "VisualTeX", _
            "The native display equation Bookmark was not persisted."
    End If

    regressionStage = "image-cross-reference-reset"
    testDocument.Content.Delete
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    Set placeholder = testDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Range:=insertionRange)
    placeholder.Width = 120
    placeholder.Height = 30
    placeholder.AlternativeText = VT_METADATA_PREFIX & "e30"
    placeholder.Title = _
        VTFormulaReference(fixtureFormulaId, "block", True)
    VTSetWordLatexPayload testDocument, fixtureFormulaId, "eF4yK3leMg"
    VTSetWordMetadataPayload _
        testDocument, fixtureFormulaId, VT_METADATA_PREFIX & "e30"
    VTSetWordFormulaFormat _
        testDocument, fixtureFormulaId, "block", True
    regressionStage = "image-cross-reference-insert-number"
    Set numberRange = VTInsertEquationNumber( _
        placeholder, fixtureFormulaId, "x^2 + y^2")
    regressionStage = "image-cross-reference-read-list"
    nativeCrossReferenceItems = _
        testDocument.GetCrossReferenceItems(wdCaptionEquation)
    If Not IsArray(nativeCrossReferenceItems) Then
        Err.Raise vbObjectError + 7488, "VisualTeX", _
            "Word did not return its native Equation cross-reference list."
    End If
    If UBound(nativeCrossReferenceItems) - _
       LBound(nativeCrossReferenceItems) + 1 <> 1 Or _
       Trim$(CStr(nativeCrossReferenceItems( _
           LBound(nativeCrossReferenceItems)))) <> "1" Then
        Err.Raise vbObjectError + 7488, "VisualTeX", _
            "Word's native Equation list is not the pure number 1."
    End If
    crossReferenceItems = _
        VTEquationNumberCrossReferenceItems(testDocument)
    If Not IsArray(crossReferenceItems) Then
        Err.Raise vbObjectError + 7488, "VisualTeX", _
            "Word did not return an Equation cross-reference list."
    End If
    itemCount = _
        UBound(crossReferenceItems) - LBound(crossReferenceItems) + 1
    If itemCount < 1 Then
        Err.Raise vbObjectError + 7489, "VisualTeX", _
            "The numbered image formula is missing from Equation cross-references."
    End If
    For itemIndex = LBound(crossReferenceItems) To UBound(crossReferenceItems)
        If Len(crossReferenceInventory) > 0 Then
            crossReferenceInventory = crossReferenceInventory & " | "
        End If
        crossReferenceInventory = crossReferenceInventory & _
            "[" & CStr(itemIndex) & "]=" & _
            Replace$(Replace$(CStr(crossReferenceItems(itemIndex)), _
                vbTab, "<TAB>"), vbCr, "<CR>")
        If Len(Trim$(CStr(crossReferenceItems(itemIndex)))) > 3 Then
            crossReferenceTextFound = True
        End If
        If Left$(Trim$(CStr(crossReferenceItems(itemIndex))), 3) = _
           "(1)" Then
            crossReferenceNumberFound = True
        End If
    Next itemIndex
    If Not crossReferenceTextFound Or Not crossReferenceNumberFound Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "The image formula Equation picker item lacks its number or preview" & _
            " [items=" & crossReferenceInventory & "]."
    End If
    If numberRange.Information(wdWithInTable) Or _
       numberRange.InlineShapes.Count <> 1 Or _
       numberRange.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7490, "VisualTeX", _
            "The image Equation did not remain in its stable paragraph layout."
    End If
    regressionStage = "image-cross-reference-assert-layout"
    VTAssertNumberedEquationLayout _
        placeholder.Range.Duplicate, placeholder.Height, fixtureFormulaId, _
        "x^2 + y^2", "image-standard"

    regressionStage = "image-cross-reference-refresh-cycle"
    testDocument.Fields.Update
    VTReconcileEquationNumbers testDocument
    VTReconcileEquationNumbers testDocument
    If Not testDocument.Bookmarks.Exists( _
       VTEquationNumberBookmarkName(fixtureFormulaId)) Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "The visible Equation number Bookmark disappeared during" & _
            " repeated field refresh."
    End If
    If Trim$(testDocument.Bookmarks( _
       VTEquationNumberBookmarkName( _
           fixtureFormulaId)).Range.Text) <> "(1)" Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "The visible Equation number Bookmark did not survive repeated" & _
            " field refresh as the complete dynamic (1)."
    End If

    regressionStage = "image-cross-reference-insert-ref"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set numberRange = VTInsertEquationNumberReferenceAtRange( _
        insertionRange, 1)
    referenceResult = Trim$(numberRange.Text)
    Set referenceField = Nothing
    For Each sequenceField In numberRange.Fields
        If sequenceField.Type = wdFieldRef Then
            Set referenceField = sequenceField
            Exit For
        End If
    Next sequenceField
    If referenceField Is Nothing Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "Word did not create the native body Equation REF field."
    End If
    If referenceResult <> "(1)" Or _
       Trim$(referenceField.Result.Text) <> "1" Or _
       InStr(1, referenceResult, "x^2 + y^2", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "The VisualTeX Equation REF is not the exact parenthesized number" & _
            " [code=" & referenceField.Code.Text & _
            "; result=" & referenceField.Result.Text & _
            "; text=" & referenceResult & "]."
    End If

    regressionStage = "image-cross-reference-insert-word-native-ref"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    insertionRange.Select
    Selection.InsertCrossReference _
        ReferenceType:=wdCaptionEquation, _
        ReferenceKind:=wdOnlyLabelAndNumber, _
        ReferenceItem:=1, _
        InsertAsHyperlink:=True, _
        IncludePosition:=False
    ' Programmatic insertion does not reliably raise WindowSelectionChange on
    ' Word for Mac. Exercise the same narrow production repair explicitly here;
    ' the physical UI regression separately validates the event route.
    VTRepairSelectedWordNativeEquationCrossReference _
        Selection.Range, privateRefRepairDiagnostic
    Set diagnosticRange = Selection.Range.Paragraphs(1).Range.Duplicate
    Set referenceField = Nothing
    For Each sequenceField In diagnosticRange.Fields
        If sequenceField.Type = wdFieldRef Then
            Set referenceField = sequenceField
            Exit For
        End If
    Next sequenceField
    If referenceField Is Nothing Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "Word did not create its built-in image Equation reference."
    End If
    referenceResult = Trim$(VTProbeParagraphText(diagnosticRange))
    If referenceResult <> "1" Or _
       Trim$(referenceField.Result.Text) <> "1" Or _
       diagnosticRange.InlineShapes.Count <> 0 Or _
       InStr(1, referenceResult, "(", vbBinaryCompare) > 0 Or _
       InStr(1, referenceResult, ")", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "Word's built-in image Equation reference is not a pure number" & _
            " [result=" & referenceField.Result.Text & _
            "; text=" & referenceResult & _
            "; images=" & CStr(diagnosticRange.InlineShapes.Count) & _
            "; repair=" & privateRefRepairDiagnostic & "]."
    End If
    builtInTargetName = VTReferenceTargetBookmarkName( _
        referenceField.Code.Text)
    If Left$(builtInTargetName, 4) <> "_Ref" Or _
       Not VTTryGetBookmarkRangeIncludingHidden( _
           testDocument, builtInTargetName, builtInTargetRange) Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "Word's repaired native Equation reference has no private target."
    End If
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, fixtureFormulaId)
    If sequenceField Is Nothing Or _
       builtInTargetRange.Start <> sequenceField.Result.Start Or _
       builtInTargetRange.End <> sequenceField.Result.End Or _
       builtInTargetRange.InlineShapes.Count <> 0 Or _
       builtInTargetRange.OMaths.Count <> 0 Or _
       VTComparableEquationNumberText(builtInTargetRange.Text) <> "1" Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "Word's repaired private _Ref is not helper-only" & _
            " [target=" & CStr(builtInTargetRange.Start) & "-" & _
                CStr(builtInTargetRange.End) & _
            "; text=" & builtInTargetRange.Text & _
            "; repair=" & privateRefRepairDiagnostic & "]."
    End If

    regressionStage = "image-cross-reference-later-number"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set referenceStabilityPlaceholder = _
        testDocument.InlineShapes.AddPicture( _
            FileName:=VTPlaceholderImagePath(), _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Range:=insertionRange)
    referenceStabilityPlaceholder.Width = 110
    referenceStabilityPlaceholder.Height = 28
    referenceStabilityPlaceholder.AlternativeText = _
        VT_METADATA_PREFIX & "e30"
    referenceStabilityPlaceholder.Title = _
        VTFormulaReference(stabilityImageFormulaId, "block", True)
    VTSetWordLatexPayload _
        testDocument, stabilityImageFormulaId, "eF4yK3leMg"
    VTSetWordMetadataPayload _
        testDocument, stabilityImageFormulaId, VT_METADATA_PREFIX & "e30"
    VTSetWordFormulaFormat _
        testDocument, stabilityImageFormulaId, "block", True
    Set numberRange = VTInsertEquationNumber( _
        referenceStabilityPlaceholder, stabilityImageFormulaId, _
        "later reference stability formula")
    VTReconcileEquationNumbers testDocument

    Set referenceField = Nothing
    For Each sequenceField In testDocument.Fields
        If sequenceField.Type = wdFieldRef And _
           Not sequenceField.Result.Information(wdWithInTable) And _
           sequenceField.Result.Paragraphs(1).Range.InlineShapes.Count = 0 And _
           sequenceField.Result.Paragraphs(1).Range.OMaths.Count = 0 Then
            Set referenceField = sequenceField
            Exit For
        End If
    Next sequenceField
    If referenceField Is Nothing Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "The earlier native body Equation REF disappeared after a later number."
    End If
    If Trim$(referenceField.Result.Text) <> "1" Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "A later numbered formula changed the earlier native Equation REF" & _
            " [result=" & referenceField.Result.Text & "]."
    End If
    nativeCrossReferenceItems = _
        testDocument.GetCrossReferenceItems(wdCaptionEquation)
    If Not IsArray(nativeCrossReferenceItems) Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "Word's native Equation list disappeared after insertion."
    End If
    If UBound(nativeCrossReferenceItems) - _
       LBound(nativeCrossReferenceItems) + 1 <> 2 Or _
       Trim$(CStr(nativeCrossReferenceItems( _
           LBound(nativeCrossReferenceItems)))) <> "1" Or _
       Trim$(CStr(nativeCrossReferenceItems( _
           LBound(nativeCrossReferenceItems) + 1))) <> "2" Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "Word's native Equation list did not preserve pure 1 and 2 items."
    End If
    crossReferenceItems = _
        VTEquationNumberCrossReferenceItems(testDocument)
    If Not IsArray(crossReferenceItems) Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "The Equation cross-reference list disappeared after" & _
            " consecutive numbered insertion."
    End If
    itemCount = _
        UBound(crossReferenceItems) - LBound(crossReferenceItems) + 1
    If itemCount <> 2 Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "The Equation cross-reference list did not retain exactly two" & _
            " numbered formulas after consecutive insertion."
    End If
    If Left$(Trim$(CStr(crossReferenceItems( _
       LBound(crossReferenceItems)))), 3) <> "(1)" Or _
       Len(Trim$(CStr(crossReferenceItems( _
           LBound(crossReferenceItems))))) <= 3 Or _
       Left$(Trim$(CStr(crossReferenceItems( _
           LBound(crossReferenceItems) + 1))), 3) <> "(2)" Or _
       Len(Trim$(CStr(crossReferenceItems( _
           LBound(crossReferenceItems) + 1)))) <= 3 Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "Consecutive Equation picker items lost their numbers or previews."
    End If

    regressionStage = "image-numbered-tall-edit"
    placeholder.LockAspectRatio = msoFalse
    placeholder.Width = 90
    placeholder.Height = 72
    placeholder.LockAspectRatio = msoTrue
    numberCreated = False
    Set numberRange = VTEnsureImageEquationNumber( _
        placeholder, placeholder.Height, fixtureFormulaId, _
        "tall fraction image", numberCreated)
    If numberCreated Then
        Err.Raise vbObjectError + 7522, "VisualTeX", _
            "Editing a tall image formula created a duplicate Equation number."
    End If
    VTAssertNumberedEquationLayout _
        placeholder.Range.Duplicate, placeholder.Height, fixtureFormulaId, _
        "tall fraction image", "image-tall-edit"

    regressionStage = "image-numbered-wide-edit"
    placeholder.LockAspectRatio = msoFalse
    placeholder.Width = 240
    placeholder.Height = 24
    placeholder.LockAspectRatio = msoTrue
    numberCreated = False
    Set numberRange = VTEnsureImageEquationNumber( _
        placeholder, placeholder.Height, fixtureFormulaId, _
        "wide aligned image", numberCreated)
    If numberCreated Then
        Err.Raise vbObjectError + 7523, "VisualTeX", _
            "Editing a wide image formula created a duplicate Equation number."
    End If
    VTAssertNumberedEquationLayout _
        placeholder.Range.Duplicate, placeholder.Height, fixtureFormulaId, _
        "wide aligned image", "image-wide-edit"

    regressionStage = "image-unnumbered-display-center"
    testDocument.Content.Delete
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    Set placeholder = testDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Range:=insertionRange)
    placeholder.Width = 160
    placeholder.Height = 40
    placeholder.Range.ParagraphFormat.Alignment = wdAlignParagraphRight
    VTNormalizeUnnumberedDisplayParagraph placeholder.Range
    Set diagnosticRange = testDocument.Range( _
        Start:=0, End:=0).Paragraphs(1).Range.Duplicate
    expectedUnnumberedPosition = VTExpectedImageFormulaPosition( _
        placeholder, CDbl(VTVisibleEquationNumberFontSize(testDocument)), _
        "The unnumbered image regression paragraph-mark alignment")
    actualUnnumberedPosition = CLng(placeholder.Range.Font.Position)
    Set paragraphMarkRange = testDocument.Range( _
        Start:=diagnosticRange.End - 1, End:=diagnosticRange.End)
    paragraphMarkPosition = CLng(paragraphMarkRange.Font.Position)
    If diagnosticRange.ParagraphFormat.Alignment <> wdAlignParagraphCenter Or _
       VTCustomTabStopCount(diagnosticRange) <> 0 Or _
       expectedUnnumberedPosition >= 0 Or _
       actualUnnumberedPosition <> expectedUnnumberedPosition Or _
       paragraphMarkPosition <> 0 Then
        Err.Raise vbObjectError + 7524, "VisualTeX", _
            "The unnumbered image display formula or its paragraph mark" & _
            " is not visually centered" & _
            " [alignment=" & _
            CStr(diagnosticRange.ParagraphFormat.Alignment) & _
            "; tabs=" & _
            CStr(diagnosticRange.ParagraphFormat.TabStops.Count) & _
            "; customTabs=" & CStr(VTCustomTabStopCount(diagnosticRange)) & _
            "; formulaPosition=" & CStr(actualUnnumberedPosition) & _
            "; expectedPosition=" & CStr(expectedUnnumberedPosition) & _
            "; paragraphMarkPosition=" & CStr(paragraphMarkPosition) & _
            "; style=" & CStr(diagnosticRange.Style) & "]."
    End If

    regressionStage = "native-numbered-create"
    testDocument.Content.Delete
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    Set equationRange = VTInsertNativeEquationAtRange( _
        insertionRange, ommlBase64, nativeDocumentPath, _
        "inline", True, False)
    nativeEquationStart = equationRange.Start
    VTSetWordOmmlPayload testDocument, nativeFormulaId, ommlBase64
    numberCreated = False
    Set numberRange = VTEnsureNativeEquationNumber( _
        equationRange, 48#, nativeFormulaId, _
        "native matrix formula", numberCreated)
    If Not numberCreated Then
        Err.Raise vbObjectError + 7525, "VisualTeX", _
            "Creating a numbered native formula did not create an Equation number."
    End If
    Set equationRange = VTResolveNativeEquationRange( _
        testDocument, nativeEquationStart, 16)
    VTAssertNumberedEquationLayout _
        equationRange, 48#, nativeFormulaId, _
        "native matrix formula", "native-numbered-create"
    displayTableAdvance = VTOMathTableVisualAdvance( _
        equationRange, wdOMathDisplay, "native-numbered-display-table")
    If displayTableAdvance <= inlineTableAdvance + 1! Then
        Err.Raise vbObjectError + 7548, "VisualTeX", _
            "Native display OMML did not occupy more vertical space than" & _
            " inline OMML in the same table layout" & _
            " [inlineAdvance=" & CStr(inlineTableAdvance) & _
            "; displayAdvance=" & CStr(displayTableAdvance) & _
            "; fontSize=" & CStr(equationRange.Font.Size) & "]."
    End If

    regressionStage = "native-display-integral-geometry"
    integralDisplayAdvance = VTNativeDocxCompactDisplayAdvance( _
        "integral", ommlBase64, integralDocumentPath, _
        "native-display-integral-geometry")
    If integralDisplayAdvance <= inlineTableAdvance + 1! Then
        Err.Raise vbObjectError + 7548, "VisualTeX", _
            "Integral display OMML did not retain display geometry" & _
            " [inlineAdvance=" & CStr(inlineTableAdvance) & _
            "; displayAdvance=" & CStr(integralDisplayAdvance) & "]."
    End If

    regressionStage = "native-display-sum-fraction-geometry"
    sumFractionDisplayAdvance = VTNativeDocxCompactDisplayAdvance( _
        "sum_fraction", ommlBase64, sumFractionDocumentPath, _
        "native-display-sum-fraction-geometry")
    If sumFractionDisplayAdvance <= inlineTableAdvance + 1! Then
        Err.Raise vbObjectError + 7548, "VisualTeX", _
            "N-ary fraction display OMML did not retain display geometry" & _
            " [inlineAdvance=" & CStr(inlineTableAdvance) & _
            "; displayAdvance=" & CStr(sumFractionDisplayAdvance) & "]."
    End If

    regressionStage = "native-numbered-create-fields"
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, nativeFormulaId)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7526, "VisualTeX", _
            "The numbered native formula has no external Equation SEQ helper."
    End If
    previousNumberText = sequenceField.Result.Text

    regressionStage = "native-numbered-edit"
    Set equationRange = VTInsertNativeEquationAtRange( _
        equationRange, ommlBase64, nativeDocumentPath, _
        "inline", True, True)
    nativeEquationStart = equationRange.Start
    numberCreated = False
    Set numberRange = VTEnsureNativeEquationNumber( _
        equationRange, 72#, nativeFormulaId, _
        "edited native matrix formula", numberCreated)
    If numberCreated Then
        Err.Raise vbObjectError + 7526, "VisualTeX", _
            "Editing a numbered native formula created a duplicate Equation number."
    End If
    Set equationRange = VTResolveNativeEquationRange( _
        testDocument, nativeEquationStart, 16)
    VTAssertNumberedEquationLayout _
        equationRange, 72#, nativeFormulaId, _
        "edited native matrix formula", "native-numbered-edit"
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, nativeFormulaId)
    If sequenceField Is Nothing Or _
       sequenceField.Result.Text <> previousNumberText Then
        Err.Raise vbObjectError + 7527, "VisualTeX", _
            "Editing a numbered native formula changed its Equation number."
    End If

    regressionStage = "image-to-native-number-preservation"
    ' Start from a genuinely new document. Deleting Content from the preceding
    ' native-number test can leave Word's external helper Frame alive and make
    ' the next picture land inside that Frame, which is not the user workflow.
    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate
    ' Match the real fresh-document failure: without a Heading 1 the chapter
    ' mode helper owns "0-1" in VT_N_, while its SEQ field result is only "1".
    ' The protected conversion must recognize that complete REF tail as managed
    ' numbering rather than reject it as surrounding body text.
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_MODE_VARIABLE, _
        VT_WORD_NUMBERING_MODE_CHAPTER
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_SEPARATOR_VARIABLE, "-"
    ' Use the same durable image identity as production insertion. The helper
    ' deliberately re-resolves that identity after Word mutates paragraph ranges.
    Set placeholder = VTRegressionCreateNumberedImage( _
        testDocument, conversionFormulaId, "eF8y", 180!, 64!)
    sourceHeightPoints = placeholder.Height
    VTAssertNumberedEquationLayout _
        placeholder.Range.Duplicate, sourceHeightPoints, conversionFormulaId, _
        "image conversion formula", "conversion-image-before"
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, conversionFormulaId)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7528, "VisualTeX", _
            "The image conversion source has no external Equation SEQ helper."
    End If
    previousNumberText = VTEquationNumberTextForFormula( _
        testDocument, conversionFormulaId)
    If VTComparableEquationNumberText(previousNumberText) <> "0-1" Then
        Err.Raise vbObjectError + 7528, "VisualTeX", _
            "The fresh chapter-mode image did not expose the expected 0-1 number."
    End If
    VTSetWordOmmlPayload testDocument, conversionFormulaId, ommlBase64
    VTWordConvertInlineShapeToNativeEquation placeholder
    Set equationRange = VTNumberedFormulaRangeForId( _
        testDocument, conversionFormulaId)
    If equationRange Is Nothing Then
        Err.Raise vbObjectError + 7528, "VisualTeX", _
            "Production image-to-native conversion lost its formula identity."
    End If
    nativeEquationStart = equationRange.Start
    numberCreated = False
    Set numberRange = VTEnsureNativeEquationNumber( _
        equationRange, sourceHeightPoints, conversionFormulaId, _
        "image conversion formula", numberCreated)
    If numberCreated Then
        Err.Raise vbObjectError + 7528, "VisualTeX", _
            "Image-to-native conversion created a duplicate Equation number."
    End If
    Set equationRange = VTResolveNativeEquationRange( _
        testDocument, nativeEquationStart, 16)
    VTAssertNumberedEquationLayout _
        equationRange, sourceHeightPoints, conversionFormulaId, _
        "image conversion formula", "conversion-native-after"
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        testDocument, conversionFormulaId)
    If sequenceField Is Nothing Or _
       VTEquationNumberTextForFormula( _
           testDocument, conversionFormulaId) <> previousNumberText Then
        Err.Raise vbObjectError + 7529, "VisualTeX", _
            "Image-to-native conversion changed its Equation number."
    End If
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_MODE_VARIABLE, _
        VT_WORD_NUMBERING_MODE_SEQUENCE
    VTSetDocumentVariable _
        testDocument, VT_WORD_NUMBERING_SEPARATOR_VARIABLE, "."

    regressionStage = "continuous-insertion-capture"
    invariantParagraphStart = equationRange.Paragraphs(1).Range.Start
    invariantSnapshot = VTNumberedEquationInvariantSnapshot( _
        testDocument, invariantParagraphStart)

    regressionStage = "continuous-insertion-image-unnumbered"
    Set newParagraph = testDocument.Paragraphs.Add
    Set insertionRange = newParagraph.Range.Duplicate
    insertionRange.Collapse wdCollapseStart
    Set placeholder = testDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), LinkToFile:=False, _
        SaveWithDocument:=True, Range:=insertionRange)
    placeholder.Width = 100
    placeholder.Height = 28
    VTNormalizeUnnumberedDisplayParagraph placeholder.Range
    VTAssertNumberedEquationInvariant _
        testDocument, invariantParagraphStart, invariantSnapshot, regressionStage

    regressionStage = "continuous-insertion-native-unnumbered"
    Set newParagraph = testDocument.Paragraphs.Add
    Set insertionRange = newParagraph.Range.Duplicate
    insertionRange.Collapse wdCollapseStart
    Set equationRange = VTInsertNativeEquationAtRange( _
        insertionRange, ommlBase64, nativeDocumentPath, _
        "block", True, False)
    VTAssertNumberedEquationInvariant _
        testDocument, invariantParagraphStart, invariantSnapshot, regressionStage

    regressionStage = "continuous-insertion-image-numbered"
    Set placeholder = VTRegressionCreateNumberedImage( _
        testDocument, stabilityImageFormulaId, "eF8y", 130!, 36!)
    VTAssertNumberedEquationInvariant _
        testDocument, invariantParagraphStart, invariantSnapshot, regressionStage

    regressionStage = "continuous-insertion-native-numbered"
    Set newParagraph = testDocument.Paragraphs.Add
    Set insertionRange = newParagraph.Range.Duplicate
    insertionRange.Collapse wdCollapseStart
    Set equationRange = VTInsertNativeEquationAtRange( _
        insertionRange, ommlBase64, nativeDocumentPath, _
        "inline", True, False)
    VTSetWordOmmlPayload testDocument, stabilityNativeFormulaId, ommlBase64
    numberCreated = False
    Set numberRange = VTEnsureNativeEquationNumber( _
        equationRange, 48#, stabilityNativeFormulaId, _
        "later native formula", numberCreated)
    If Not numberCreated Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The later numbered native formula did not create its own number."
    End If
    VTAssertNumberedEquationInvariant _
        testDocument, invariantParagraphStart, invariantSnapshot, regressionStage

    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    VTWriteTextAtomic _
        fixtureRoot & "/word-native-regression-result.txt", _
        "PASS" & vbLf & "crossReferenceItems=" & CStr(itemCount) & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic _
        fixtureRoot & "/word-native-regression-result.txt", _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), vbLf, " ") & vbLf
    ' Preserve the failed document for popup/process/range inspection.
    If Not testDocument Is Nothing Then
        VTWriteTextAtomic _
            fixtureRoot & "/word-native-regression-failure-openxml.xml", _
            testDocument.Content.WordOpenXML
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word native regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Private Function VTOMathTableVisualAdvance( _
    ByVal equationRange As Range, _
    ByVal expectedType As Long, _
    ByVal assertionName As String) As Single

    Dim sourceDocument As Document
    Dim measurementDocument As Document
    Dim layoutTable As Table
    Dim insertionRange As Range
    Dim centerRange As Range
    Dim formulaInsert As Range
    Dim formulaParagraph As Range
    Dim afterTableParagraph As Range
    Dim nativeEquation As OMath
    Dim cellXml As String
    Dim formulaY As Single
    Dim afterTableY As Single
    Dim measurementErrorNumber As Long
    Dim measurementErrorDescription As String

    If equationRange Is Nothing Or equationRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7548, "VisualTeX", _
            assertionName & ": table measurement has no unique OMath."
    End If
    Set sourceDocument = equationRange.Document
    On Error GoTo MeasurementFailed
    Set measurementDocument = Documents.Add(Visible:=True)
    measurementDocument.ActiveWindow.View.Type = wdPrintView
    Set insertionRange = measurementDocument.Range(Start:=0, End:=0)
    If expectedType = wdOMathInline Then
        Set layoutTable = measurementDocument.Tables.Add( _
            Range:=insertionRange, NumRows:=1, NumColumns:=3)
        VTConfigureNumberedDisplayTable layoutTable
        Set centerRange = layoutTable.Cell(1, 2).Range.Duplicate
        centerRange.End = centerRange.End - 1
        centerRange.Text = "AZ"
        Set formulaInsert = measurementDocument.Range( _
            Start:=layoutTable.Cell(1, 2).Range.Start + 1, _
            End:=layoutTable.Cell(1, 2).Range.Start + 1)
        formulaInsert.FormattedText = equationRange.FormattedText
    Else
        insertionRange.FormattedText = equationRange.FormattedText
        If measurementDocument.OMaths.Count <> 1 Then
            Err.Raise vbObjectError + 7548, "VisualTeX", _
                assertionName & ": display probe has no unique OMath."
        End If
        Set layoutTable = VTWrapNativeDisplayParagraphInTable( _
            VTNativeMathRoot(measurementDocument.OMaths(1)).Range.Duplicate)
    End If
    If layoutTable.Cell(1, 2).Range.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7548, "VisualTeX", _
            assertionName & ": copied table cell has no unique OMath."
    End If
    Set nativeEquation = layoutTable.Cell(1, 2).Range.OMaths(1)
    nativeEquation.Type = expectedType
    If expectedType = wdOMathDisplay Then
        nativeEquation.Justification = wdOMathJcCenter
    End If
    nativeEquation.BuildUp
    nativeEquation.Range.Font.Position = 0
    layoutTable.Cell(1, 2).Range.ParagraphFormat.Alignment = _
        wdAlignParagraphCenter
    If nativeEquation.Type <> expectedType Then
        Err.Raise vbObjectError + 7548, "VisualTeX", _
            assertionName & ": Word did not preserve the requested OMath type" & _
            " [actual=" & CStr(nativeEquation.Type) & _
            "; expected=" & CStr(expectedType) & "]."
    End If
    If expectedType = wdOMathDisplay Then
        If layoutTable.Cell(1, 2).Range.Paragraphs.Count <> 2 Then
            Err.Raise vbObjectError + 7548, "VisualTeX", _
                assertionName & _
                ": display table lost its required compact tail paragraph."
        End If
        cellXml = layoutTable.Cell(1, 2).Range.WordOpenXML
        If InStr(1, cellXml, "<m:oMathPara", vbBinaryCompare) = 0 Then
            Err.Raise vbObjectError + 7548, "VisualTeX", _
                assertionName & ": display table has no m:oMathPara."
        End If
        If layoutTable.Cell(1, 2).Range.Paragraphs(2).Range.Font.Size > 1.5 Or _
           layoutTable.Cell(1, 2).Range.Paragraphs(2).Range.ParagraphFormat.LineSpacing > 1.5 Then
            Err.Raise vbObjectError + 7548, "VisualTeX", _
                assertionName & ": display table tail was not compacted."
        End If
    End If

    measurementDocument.Repaginate
    Set formulaParagraph = _
        layoutTable.Cell(1, 2).Range.Paragraphs(1).Range.Duplicate
    Set afterTableParagraph = measurementDocument.Range( _
        Start:=layoutTable.Range.End, _
        End:=layoutTable.Range.End).Paragraphs(1).Range.Duplicate
    formulaY = CSng(formulaParagraph.Information( _
        wdVerticalPositionRelativeToPage))
    afterTableY = CSng(afterTableParagraph.Information( _
        wdVerticalPositionRelativeToPage))
    If formulaY < 0! Or afterTableY < 0! Or _
       afterTableY - formulaY <= 0! Then
        Err.Raise vbObjectError + 7548, "VisualTeX", _
            assertionName & ": Word did not expose a measurable table advance" & _
            " [formulaY=" & CStr(formulaY) & _
            "; afterTableY=" & CStr(afterTableY) & "]."
    End If
    VTOMathTableVisualAdvance = afterTableY - formulaY

    measurementDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set measurementDocument = Nothing
    sourceDocument.Activate
    Exit Function

MeasurementFailed:
    measurementErrorNumber = Err.Number
    measurementErrorDescription = Err.Description
    On Error Resume Next
    If Not measurementDocument Is Nothing Then
        measurementDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    On Error GoTo 0
    Err.Raise vbObjectError + 7548, "VisualTeX", _
        assertionName & ": table measurement failed" & _
        " [error=" & CStr(measurementErrorNumber) & _
        "; description=" & measurementErrorDescription & "]."
End Function

Private Function VTNativeDocxCompactDisplayAdvance( _
    ByVal fixtureName As String, _
    ByVal validationOmmlBase64 As String, _
    ByVal nativeDocumentPath As String, _
    ByVal assertionName As String) As Double

    Dim owningDocument As Document
    Dim resultLine As String
    Dim resultFields As Variant
    Dim formulaAdvance As Double
    Dim anchorSpan As Double
    Dim fontSize As Double

    On Error Resume Next
    Set owningDocument = ActiveDocument
    On Error GoTo 0
    resultLine = VTProbeOneDisplayStrategy( _
        fixtureName, "formatted-paragraph-compact-tail", _
        validationOmmlBase64, nativeDocumentPath)
    If Not owningDocument Is Nothing Then owningDocument.Activate
    resultFields = Split(resultLine, "|")
    If UBound(resultFields) < 15 Then
        Err.Raise vbObjectError + 7548, "VisualTeX", _
            assertionName & ": compact display probe returned malformed data" & _
            " [result=" & resultLine & "]."
    End If
    If CStr(resultFields(2)) <> "OK" Or _
       CLng(Val(CStr(resultFields(5)))) <> wdOMathDisplay Or _
       CLng(Val(CStr(resultFields(6)))) < 1 Or _
       CLng(Val(CStr(resultFields(7)))) <> wdOMathDisplay Or _
       CLng(Val(CStr(resultFields(8)))) < 1 Or _
       CLng(Val(CStr(resultFields(9)))) <> 1 Or _
       CLng(Val(CStr(resultFields(10)))) <> 2 Then
        Err.Raise vbObjectError + 7548, "VisualTeX", _
            assertionName & ": compact display probe did not preserve" & _
            " one two-paragraph m:oMathPara cell" & _
            " [result=" & resultLine & "]."
    End If
    formulaAdvance = Val(CStr(resultFields(11)))
    anchorSpan = Val(CStr(resultFields(12)))
    fontSize = Val(CStr(resultFields(13)))
    If formulaAdvance <= 0# Or anchorSpan <= 0# Or _
       fontSize <= 0# Or fontSize > 72# Then
        Err.Raise vbObjectError + 7548, "VisualTeX", _
            assertionName & ": compact display probe geometry is invalid" & _
            " [result=" & resultLine & "]."
    End If
    VTNativeDocxCompactDisplayAdvance = formulaAdvance
End Function

Public Sub VisualTeX_DisableWordEventsForRegression()
    Set VT_WORD_EVENT_SINK = Nothing
End Sub

Private Function VTBodyReferenceResultForTarget( _
    ByVal documentObject As Document, _
    ByVal targetBookmarkName As String) As String

    Dim candidateField As Field
    Dim matchCount As Long
    Dim candidateResult As String
    Dim commonResult As String

    If documentObject Is Nothing Or Len(targetBookmarkName) = 0 Then
        Exit Function
    End If
    For Each candidateField In documentObject.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) And _
           StrComp(VTReferenceTargetBookmarkName( _
               candidateField.Code.Text), targetBookmarkName, _
               vbTextCompare) = 0 Then
            candidateResult = Trim$(candidateField.Result.Text)
            matchCount = matchCount + 1
            If matchCount = 1 Then
                commonResult = candidateResult
            ElseIf candidateResult <> commonResult Then
                commonResult = ""
                Exit For
            End If
        End If
    Next candidateField
    If matchCount > 0 Then VTBodyReferenceResultForTarget = commonResult
End Function

Private Sub VTVerifyNumberedFormulaIntegrity( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal expectedOrdinal As Long)

    Dim formulaRange As Range
    Dim layoutTable As Table

    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7561, "VisualTeX", _
            "The numbered formula verification target is missing."
    End If
    If formulaRange.Information(wdWithInTable) Then
        Set layoutTable = formulaRange.Tables(1)
        VTVerifyEquationNumberFieldIntegrity _
            layoutTable, formulaId, expectedOrdinal
    Else
        VTVerifyParagraphEquationNumberIntegrity _
            formulaRange, formulaId, expectedOrdinal
    End If
End Sub

Private Function VTMacroButtonResultDiagnostic( _
    ByVal macroField As Field) As String

    Dim resultText As String
    Dim firstCode As Long
    Dim secondCode As Long

    On Error Resume Next
    If macroField Is Nothing Then
        VTMacroButtonResultDiagnostic = "field=nothing"
        Exit Function
    End If
    resultText = macroField.Result.Text
    If Len(resultText) >= 1 Then firstCode = AscW(Left$(resultText, 1))
    If Len(resultText) >= 2 Then secondCode = AscW(Mid$(resultText, 2, 1))
    VTMacroButtonResultDiagnostic = _
        "type=" & CStr(macroField.Type) & _
        " shapes=" & CStr(macroField.Result.InlineShapes.Count) & _
        " omaths=" & CStr(macroField.Result.OMaths.Count) & _
        " textLen=" & CStr(Len(resultText)) & _
        " firstCode=" & CStr(firstCode) & _
        " secondCode=" & CStr(secondCode)
    On Error GoTo 0
End Function

Public Sub VisualTeX_RefreshImageMacroButtonsForRegression()
    Dim candidateField As Field
    Dim macroButtonCount As Long

    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7593, "VisualTeX regression", _
            "There is no active Word document for MacroButton verification."
    End If
    For Each candidateField In ActiveDocument.Fields
        If VTIsVisualTeXImageMacroButtonField(candidateField) Then
            macroButtonCount = macroButtonCount + 1
            If Not candidateField.Locked Then
                Err.Raise vbObjectError + 7593, "VisualTeX regression", _
                    "A committed VisualTeX image MacroButton is not locked."
            End If
            ' Exercise Word's normal field-refresh request while the field is
            ' locked. Some builds return error 4605 and others report a no-op;
            ' either result is acceptable only if the pure picture survives.
            On Error Resume Next
            Err.Clear
            candidateField.Update
            Err.Clear
            On Error GoTo 0
            If Not VTMacroButtonResultHasOnlyImage(candidateField) Then
                Err.Raise vbObjectError + 7593, "VisualTeX regression", _
                    "A Word field refresh replaced the VisualTeX formula picture: " & _
                    VTMacroButtonResultDiagnostic(candidateField)
            End If
        End If
    Next candidateField
    If macroButtonCount = 0 Then
        Err.Raise vbObjectError + 7593, "VisualTeX regression", _
            "The test document contains no VisualTeX image MacroButton."
    End If
End Sub

Public Sub VisualTeX_WriteSelectedPictureScreenBoundsRegression()
    Dim targetRange As Range
    Dim screenLeft As Long
    Dim screenTop As Long
    Dim screenWidth As Long
    Dim screenHeight As Long
    Dim statusPath As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    statusPath = VTApplicationSupportRoot() & _
        "/physical-double-click-screen-bounds.txt"
    On Error GoTo Failed
    If Documents.Count = 0 Or Selection Is Nothing Then
        Err.Raise vbObjectError + 7593, "VisualTeX regression", _
            "There is no selected Word picture for physical double-click bounds."
    End If
    If Selection.InlineShapes.Count = 1 Then
        Set targetRange = Selection.InlineShapes(1).Range.Duplicate
    ElseIf Selection.ShapeRange.Count = 1 Then
        Set targetRange = Selection.ShapeRange(1).Anchor.Duplicate
    End If
    If targetRange Is Nothing Then
        Err.Raise vbObjectError + 7593, "VisualTeX regression", _
            "The selected physical double-click target is not a Word picture."
    End If
    ActiveWindow.GetPoint _
        screenLeft, screenTop, screenWidth, screenHeight, targetRange
    If screenWidth <= 0 Or screenHeight <= 0 Then
        Err.Raise vbObjectError + 7593, "VisualTeX regression", _
            "Word returned invalid picture screen bounds."
    End If
    VTWriteTextAtomic statusPath, _
        "PASS|" & CStr(screenLeft) & "|" & CStr(screenTop) & "|" & _
        CStr(screenWidth) & "|" & CStr(screenHeight) & vbLf
    Exit Sub

Failed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic statusPath, _
        "FAIL|" & CStr(regressionErrorNumber) & "|" & _
        Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
            vbLf, " ") & vbLf
    On Error GoTo 0
End Sub

Public Sub VisualTeX_WriteSelectedScreenBoundsForRegression()
    Dim formulaShape As InlineShape
    Dim nativeBookmark As Bookmark
    Dim targetRange As Range
    Dim screenLeft As Long
    Dim screenTop As Long
    Dim screenWidth As Long
    Dim screenHeight As Long
    Dim statusPath As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    statusPath = VTApplicationSupportRoot() & _
        "/physical-double-click-screen-bounds.txt"
    On Error GoTo Failed
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7593, "VisualTeX regression", _
            "There is no active Word document for physical double-click bounds."
    End If
    Set formulaShape = VTVisualTeXInlineShapeAtSelection(Selection)
    If Not formulaShape Is Nothing Then
        Set targetRange = formulaShape.Range.Duplicate
    Else
        Set nativeBookmark = _
            VTFindNativeFormulaBookmark(Selection.Range, False)
        If Not nativeBookmark Is Nothing Then
            Set targetRange = nativeBookmark.Range.Duplicate
        End If
    End If
    If targetRange Is Nothing Then
        Err.Raise vbObjectError + 7593, "VisualTeX regression", _
            "The selected physical double-click target is not a VisualTeX formula."
    End If
    ActiveWindow.GetPoint _
        screenLeft, screenTop, screenWidth, screenHeight, targetRange
    If screenWidth <= 0 Or screenHeight <= 0 Then
        Err.Raise vbObjectError + 7593, "VisualTeX regression", _
            "Word returned invalid physical double-click screen bounds."
    End If
    VTWriteTextAtomic statusPath, _
        "PASS|" & CStr(screenLeft) & "|" & CStr(screenTop) & "|" & _
        CStr(screenWidth) & "|" & CStr(screenHeight) & vbLf
    Exit Sub

Failed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic statusPath, _
        "FAIL|" & CStr(regressionErrorNumber) & "|" & _
        Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
            vbLf, " ") & vbLf
    On Error GoTo 0
End Sub

Public Sub VisualTeX_RunActiveDocumentImageSchemaMigrationRegression()
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7593, "VisualTeX regression", _
            "The image schema migration regression has no active document."
    End If
    VTMigrateDocumentImageMacroButtons ActiveDocument
End Sub

Private Sub VTCreatePictureRoutingRegressionFixture( _
    ByVal fixtureKind As String)

    Const damagedFormulaId As String = _
        "11111111-1111-4111-8111-111111111111"
    Const fixturePath As String = _
        "/Users/lpj/Library/Group Containers/UBF8T346G9.Office/" & _
        "VisualTeX/Scratch/picture-routing-fixture.png"

    Dim insertionRange As Range
    Dim inlinePicture As InlineShape
    Dim floatingPicture As Shape

    If Documents.Count = 0 Then Documents.Add
    Set insertionRange = ActiveDocument.Content
    insertionRange.Collapse wdCollapseEnd
    Set inlinePicture = ActiveDocument.InlineShapes.AddPicture( _
        FileName:=fixturePath, _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    inlinePicture.LockAspectRatio = msoFalse
    inlinePicture.Width = 72
    inlinePicture.Height = 36
    inlinePicture.LockAspectRatio = msoTrue

    Select Case fixtureKind
        Case "forged-prefix"
            inlinePicture.Title = _
                VT_FORMULA_REF_PREFIX & "not-a-uuid:inline:0"
            inlinePicture.Range.Select
        Case "damaged-metadata"
            inlinePicture.Title = _
                VTFormulaReference(damagedFormulaId, "inline", False)
            inlinePicture.AlternativeText = ""
            inlinePicture.Range.Select
        Case "ordinary-floating"
            Set floatingPicture = inlinePicture.ConvertToShape
            floatingPicture.RelativeHorizontalPosition = _
                wdRelativeHorizontalPositionPage
            floatingPicture.RelativeVerticalPosition = _
                wdRelativeVerticalPositionPage
            floatingPicture.Left = 180
            floatingPicture.Top = 180
            floatingPicture.Width = 72
            floatingPicture.Height = 36
            floatingPicture.Select
        Case Else
            inlinePicture.Title = ""
            inlinePicture.AlternativeText = ""
            inlinePicture.Range.Select
    End Select
End Sub

Public Sub VisualTeX_CreateOrdinaryInlinePictureRegression()
    VTCreatePictureRoutingRegressionFixture "ordinary-inline"
End Sub

Public Sub VisualTeX_CreateOrdinaryFloatingPictureRegression()
    VTCreatePictureRoutingRegressionFixture "ordinary-floating"
End Sub

Public Sub VisualTeX_CreateForgedPrefixPictureRegression()
    VTCreatePictureRoutingRegressionFixture "forged-prefix"
End Sub

Public Sub VisualTeX_CreateDamagedMetadataPictureRegression()
    VTCreatePictureRoutingRegressionFixture "damaged-metadata"
End Sub

Public Sub VisualTeX_RegressionConvertSelectedImageToOmml()
    VisualTeX_ConvertSelectedToNativeEquation
End Sub

Public Sub VisualTeX_RegressionConvertSelectedOmmlToImage()
    VisualTeX_ConvertSelectedToImageFormula
End Sub

Private Function VTCountManagedEquationSequences( _
    ByVal documentObject As Document) As Long

    Dim candidate As Field
    Dim math As OMath
    Dim imageCurrent As Field
    Dim imageIncrement As Field
    Dim sequenceBookmarkName As String

    If documentObject Is Nothing Then Exit Function
    For Each math In documentObject.OMaths
        If Not VTPureNativeNumberField(math.Range) Is Nothing Then
            VTCountManagedEquationSequences = _
                VTCountManagedEquationSequences + 1
        End If
    Next math
    For Each candidate In documentObject.Fields
        If VTIsImagePlaceRef(candidate) Then
            Set imageCurrent = VTImageSequenceField(candidate, False)
            Set imageIncrement = VTImageSequenceField(candidate, True)
            If Not imageCurrent Is Nothing And _
               Not imageIncrement Is Nothing Then
                VTCountManagedEquationSequences = _
                    VTCountManagedEquationSequences + 1
            End If
        ElseIf VTIsNativeEquationSequenceField( _
           candidate, VTNativeEquationLabelName()) Then
            sequenceBookmarkName = VTSequenceBookmarkNameForField( _
                documentObject, candidate)
            If Left$(sequenceBookmarkName, _
               Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)) = _
               VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX Then
                VTCountManagedEquationSequences = _
                    VTCountManagedEquationSequences + 1
            End If
        End If
    Next candidate
End Function

Public Sub VTRegressionProfileActiveEquationRefresh()
    Dim report As String
    Dim passIndex As Long
    Dim updated As Long
    Dim beforeText As String
    Dim beforeFields As Long
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    beforeText = ActiveDocument.Content.Text
    beforeFields = ActiveDocument.Fields.Count
    For passIndex = 1 To 3
        VTWordPerformanceStart
        VT_WORD_PERFORMANCE_ACTIVE = True
        updated = VTUpdateEquationNumbersCore(ActiveDocument, True)
        If ActiveDocument.Content.Text <> beforeText Or _
           ActiveDocument.Fields.Count <> beforeFields Then
            Err.Raise vbObjectError + 7598, "VisualTeX regression", _
                "An unchanged refresh modified text or field count."
        End If
        If VTWordInternalMutationActive() Or VT_WORD_NUMBER_UPDATE_RUNNING Or _
           Not Application.ScreenUpdating Then
            Err.Raise vbObjectError + 7598, "VisualTeX regression", _
                "An unchanged refresh left Word in an operation state."
        End If
        report = report & "pass=" & CStr(passIndex) & _
            ";updated=" & CStr(updated) & vbLf & VT_WORD_PERFORMANCE_BUFFER
    Next passIndex
    GoTo Finished
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
Finished:
    VT_WORD_PERFORMANCE_ACTIVE = False
    report = report & "errorNumber=" & CStr(failureNumber) & vbLf & _
        "errorDescription=" & failureDescription & vbLf
    VTWriteTextAtomic VTApplicationSupportRoot() & _
        "/numbering-refresh-regression.txt", report
End Sub

Private Function VTNativeEquationReferenceItemForFormula( _
    ByVal documentObject As Document, _
    ByVal formulaId As String) As Long

    Dim targetField As Field
    Dim candidateField As Field
    Dim sequenceBookmarkName As String
    Dim nativeItemIndex As Long
    Dim targetStart As Long

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then
        Exit Function
    End If
    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    Set targetField = VTEquationSequenceFieldForBookmark( _
        documentObject, sequenceBookmarkName)
    If targetField Is Nothing Then Exit Function
    targetStart = VTEquationFieldStart(targetField)

    For Each candidateField In documentObject.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            nativeItemIndex = nativeItemIndex + 1
            If VTEquationFieldStart(candidateField) = targetStart Then
                VTNativeEquationReferenceItemForFormula = nativeItemIndex
                Exit Function
            End If
        End If
    Next candidateField
End Function

Public Sub VisualTeX_SetDocumentImportStatusEncodingRegression()
    Dim statusText As String

    statusText = VTDocumentImportStatusPrefix("inserting") & "11/35"
    Application.StatusBar = statusText
    VTWriteTextAtomic _
        VTApplicationSupportRoot() & _
            "/document-import-status-encoding-regression.txt", _
        statusText
End Sub

Public Sub VisualTeX_ClearDocumentImportStatusEncodingRegression()
    Application.StatusBar = ""
End Sub

Public Sub VisualTeX_ApplyPendingResultForRegression()
    Dim sessionId As String
    Dim statusPath As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo Failed
    sessionId = VTReadActiveSessionId(VT_WORD_HOST)
    statusPath = VTSessionDirectory(sessionId) & "/word-callback-status.txt"
    VisualTeX_ApplyPendingResult
    VTWriteTextAtomic statusPath, "PASS" & vbLf
    Exit Sub

Failed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    If Len(statusPath) = 0 And Len(sessionId) > 0 Then
        statusPath = VTSessionDirectory(sessionId) & _
            "/word-callback-status.txt"
    End If
    If Len(statusPath) > 0 Then
        VTWriteTextAtomic statusPath, _
            "FAIL" & vbLf & _
            "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
            "errorDescription=" & Replace$(Replace$( _
                regressionErrorDescription, vbCr, " "), vbLf, " ") & vbLf
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Function VTResolveImageFormulaInParagraph( _
    ByVal documentObject As Document, _
    ByVal paragraphStart As Long) As InlineShape

    Dim paragraphRange As Range

    If documentObject Is Nothing Or paragraphStart < 0 Or _
       paragraphStart >= documentObject.Content.End Then
        Err.Raise vbObjectError + 7538, "VisualTeX", _
            "The numbered image paragraph anchor is invalid."
    End If
    Set paragraphRange = documentObject.Range( _
        Start:=paragraphStart, End:=paragraphStart).Paragraphs(1).Range.Duplicate
    If paragraphRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7538, "VisualTeX", _
            "Word could not re-resolve exactly one numbered formula image."
    End If
    Set VTResolveImageFormulaInParagraph = paragraphRange.InlineShapes(1)
End Function

Private Function VTPrependCenterTabPreservingImage( _
    ByVal formulaRange As Range) As Range

    Dim documentObject As Document
    Dim backupDocument As Document
    Dim backupRange As Range
    Dim insertionRange As Range
    Dim restoredRange As Range
    Dim containerRange As Range
    Dim sourceShape As InlineShape
    Dim restoredShape As InlineShape
    Dim containingField As Field
    Dim formulaStart As Long
    Dim sourceAlternativeText As String
    Dim sourceTitle As String
    Dim sourceFormulaId As String
    Dim sourceDisplayMode As String
    Dim sourceNumbered As Boolean
    Dim operationErrorNumber As Long
    Dim operationErrorDescription As String

    If formulaRange Is Nothing Or formulaRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7542, "VisualTeX", _
            "The numbered formula image backup target is invalid."
    End If
    Set documentObject = formulaRange.Document
    Set sourceShape = formulaRange.InlineShapes(1)
    ' Capture metadata before resolving the MacroButton container. Word for Mac
    ' can invalidate the InlineShape wrapper while its containing Field/Range is
    ' being resolved, even though the picture remains in the document.
    sourceAlternativeText = sourceShape.AlternativeText
    sourceTitle = sourceShape.Title
    If Not VTTryParseFormulaReference( _
       sourceTitle, sourceFormulaId, sourceDisplayMode, sourceNumbered) Then
        Err.Raise vbObjectError + 7542, "VisualTeX", _
            "The numbered formula image identity is invalid."
    End If
    Set containingField = VTContainingFieldForInlineShape(sourceShape)
    Set containerRange = _
        VTVisualTeXImageContainerRange(sourceShape)
    Set formulaRange = containerRange.Duplicate
    formulaStart = formulaRange.Start
    On Error GoTo RestoreFailed

    ' Current VisualTeX images are plain InlineShapes. Inserting one ordinary
    ' tab immediately before a plain picture does not mutate the picture itself,
    ' so avoid the legacy hidden-document FormattedText round-trip. That older
    ' path opened one temporary Word document per numbering operation and, on
    ' Word for Mac, those documents could remain resident even after Close,
    ' causing sequential insertion latency to grow with every formula.
    If containingField Is Nothing Then
        Set insertionRange = documentObject.Range( _
            Start:=formulaStart, End:=formulaStart)
        insertionRange.Text = vbTab
        Set restoredShape = VTResolveImageFormulaByIdentity( _
            documentObject, sourceFormulaId, sourceAlternativeText, _
            sourceTitle, formulaStart)
        If restoredShape.AlternativeText <> sourceAlternativeText Or _
           restoredShape.Title <> sourceTitle Or _
           Not VTContainingFieldForInlineShape(restoredShape) Is Nothing Then
            Err.Raise vbObjectError + 7542, "VisualTeX", _
                "Word changed the VisualTeX image while inserting the center tab."
        End If
        Set VTPrependCenterTabPreservingImage = _
            restoredShape.Range.Duplicate
        Exit Function
    End If

    Set sourceShape = Nothing
    Set backupDocument = Documents.Add(Visible:=False)
    Set backupRange = backupDocument.Content
    backupRange.Collapse wdCollapseStart
    backupRange.FormattedText = formulaRange.FormattedText
    If backupDocument.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7542, "VisualTeX", _
            "Word could not back up the numbered formula image."
    End If
    Set backupRange = backupDocument.InlineShapes(1).Range.Duplicate

    documentObject.Activate
    formulaRange.Delete
    Set insertionRange = documentObject.Range( _
        Start:=formulaStart, End:=formulaStart)
    insertionRange.InsertBefore vbTab
    Set insertionRange = documentObject.Range( _
        Start:=formulaStart + 1, End:=formulaStart + 1)
    insertionRange.FormattedText = backupRange.FormattedText
    Set restoredRange = documentObject.Range( _
        Start:=formulaStart + 1, _
        End:=documentObject.Range( _
            formulaStart + 1, formulaStart + 1).Paragraphs(1).Range.End)
    If restoredRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7542, "VisualTeX", _
            "Word could not restore the numbered formula image after the center tab."
    End If
    Set restoredShape = restoredRange.InlineShapes(1)
    restoredShape.AlternativeText = sourceAlternativeText
    restoredShape.Title = sourceTitle
    Set restoredShape = VTEnsureVisualTeXImageMacroButton(restoredShape)
    ' Field creation and result replacement are structural mutations. Resolve
    ' the final InlineShape again from its paragraph before reading metadata.
    Set restoredShape = VTResolveImageFormulaByIdentity( _
        documentObject, sourceFormulaId, sourceAlternativeText, _
        sourceTitle, formulaStart + 1)
    If restoredShape.AlternativeText <> _
       sourceAlternativeText Or _
       restoredShape.Title <> sourceTitle Then
        Err.Raise vbObjectError + 7542, "VisualTeX", _
            "Word did not preserve the VisualTeX image identity after layout."
    End If
    Set VTPrependCenterTabPreservingImage = _
        restoredShape.Range.Duplicate
    backupDocument.Close SaveChanges:=wdDoNotSaveChanges
    Exit Function

RestoreFailed:
    operationErrorNumber = Err.Number
    operationErrorDescription = Err.Description
    On Error Resume Next
    If Not backupDocument Is Nothing Then
        backupDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise operationErrorNumber, "VisualTeX Equation numbering", _
        "VTPrependCenterTabPreservingImage: " & operationErrorDescription
End Function

Private Function VTTryCalculateCurrentNumberedImagePaintedCenterDeltaPoints( _
    ByVal formulaShape As InlineShape, _
    ByVal numberRange As Range, _
    ByRef centerDeltaPoints As Double) As Boolean

    Dim formulaHeight As Double
    Dim numberSize As Double
    Dim inkCenterYRatio As Double
    Dim appliedPosition As Double
    Dim positionReadable As Boolean
    Dim formulaInkCenterAboveBaseline As Double
    Dim numberInkCenterAboveBaseline As Double

    If formulaShape Is Nothing Or numberRange Is Nothing Then Exit Function
    formulaHeight = CDbl(formulaShape.Height)
    numberSize = CDbl(numberRange.Font.Size)
    If formulaHeight <= 0# Or formulaHeight > 10000# Or _
       numberSize <= 0# Or _
       numberSize > VT_WORD_MAX_FORMULA_FONT_SIZE_PT Then Exit Function

    On Error Resume Next
    Err.Clear
    appliedPosition = CDbl(formulaShape.Range.Font.Position)
    positionReadable = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    If Not positionReadable Or appliedPosition = wdUndefined Or _
       appliedPosition < -256# Or appliedPosition > 256# Then Exit Function

    inkCenterYRatio = VTNumberedImageInkCenterYRatio(formulaShape)
    formulaInkCenterAboveBaseline = _
        formulaHeight * (1# - inkCenterYRatio) + appliedPosition
    numberInkCenterAboveBaseline = _
        numberSize * _
        VT_WORD_EQUATION_NUMBER_INK_CENTER_ABOVE_BASELINE_RATIO
    Select Case formulaShape.Range.ParagraphFormat.BaseLineAlignment
        Case wdBaselineAlignTop
            numberInkCenterAboveBaseline = formulaHeight - numberSize * _
                (0.463134765625# + 0.9501953125# - 0.70556640625#)
        Case wdBaselineAlignCenter
            numberInkCenterAboveBaseline = formulaHeight / 2# + numberSize * _
                (0.242431640625# - (0.9501953125# - 0.22216796875#) / 2#)
        Case wdBaselineAlignFarEast50
            numberInkCenterAboveBaseline = numberSize * (0.9501953125# - 0.22216796875#) / 2#
    End Select
    ' Positive means the picture's painted centre is lower on the page than the
    ' number's painted centre. Whole-point Font.Position quantisation guarantees
    ' that a correctly rounded result stays within half a point.
    centerDeltaPoints = _
        numberInkCenterAboveBaseline - formulaInkCenterAboveBaseline
    VTTryCalculateCurrentNumberedImagePaintedCenterDeltaPoints = True
End Function

Private Function VTNativeEquationVisibleHorizontalBounds( _
    ByVal formulaContentRange As Range, _
    ByRef leftPosition As Single, _
    ByRef rightPosition As Single) As Boolean

    Dim characterIndex As Long
    Dim candidateRange As Range
    Dim startProbe As Range
    Dim endProbe As Range
    Dim candidateStart As Single
    Dim candidateEnd As Single
    Dim foundVisibleCharacter As Boolean

    If formulaContentRange Is Nothing Then Exit Function
    leftPosition = 1000000!
    rightPosition = -1!

    For characterIndex = 1 To formulaContentRange.Characters.Count
        Set candidateRange = _
            formulaContentRange.Characters(characterIndex).Duplicate
        If VTWordRangeHasMeaningfulText(candidateRange) Then
            Set startProbe = candidateRange.Duplicate
            startProbe.Collapse wdCollapseStart
            Set endProbe = candidateRange.Duplicate
            endProbe.Collapse wdCollapseEnd
            candidateStart = CSng(startProbe.Information( _
                wdHorizontalPositionRelativeToTextBoundary))
            candidateEnd = CSng(endProbe.Information( _
                wdHorizontalPositionRelativeToTextBoundary))
            If candidateStart >= 0! And candidateEnd >= 0! Then
                If candidateStart < leftPosition Then _
                    leftPosition = candidateStart
                If candidateEnd < leftPosition Then _
                    leftPosition = candidateEnd
                If candidateStart > rightPosition Then _
                    rightPosition = candidateStart
                If candidateEnd > rightPosition Then _
                    rightPosition = candidateEnd
                foundVisibleCharacter = True
            End If
        End If
    Next characterIndex

    VTNativeEquationVisibleHorizontalBounds = _
        foundVisibleCharacter And rightPosition > leftPosition
End Function

Private Sub VTAssertNativeEquationArrayLayout( _
    ByVal formulaRange As Range, _
    ByVal formulaId As String, _
    ByVal expectedCaptionText As String, _
    ByVal assertionName As String)

    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim helperParagraph As Range
    Dim formulaContentRange As Range
    Dim markerRange As Range
    Dim sequenceField As Field
    Dim visibleNumberField As Field
    Dim candidateField As Field
    Dim numberRange As Range
    Dim expectedNumberRange As Range
    Dim sequenceRange As Range
    Dim captionRange As Range
    Dim nativeBookmarkRange As Range
    Dim numberEndProbe As Range
    Dim formulaProbe As Range
    Dim numberProbe As Range
    Dim nativeItems As Variant
    Dim nativeItemIndex As Long
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim textWidth As Single
    Dim formulaStartPosition As Single
    Dim formulaEndPosition As Single
    Dim formulaCenterPosition As Single
    Dim numberEndPosition As Single
    Dim formulaY As Single
    Dim numberY As Single
    Dim numberBookmarkName As String
    Dim sequenceBookmarkName As String
    Dim captionBookmarkName As String
    Dim nativeBookmarkName As String
    Dim expectedNumberText As String

    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & ": native display Equation is missing."
    End If
    If formulaRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & ": native display Equation is ambiguous."
    End If
    Set documentObject = formulaRange.Document
    Set formulaRange = VTNativeMathRoot(formulaRange.OMaths(1)).Range.Duplicate
    Set paragraphRange = VTWordParagraphContainingFormula(formulaRange)
    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    Set sequenceField = VTEquationSequenceFieldForBookmark( _
        documentObject, sequenceBookmarkName)
    Set visibleNumberField = VTNativeEquationArrayReferenceField( _
        formulaRange, formulaId)
    Set markerRange = VTNativeEquationArrayMarkerRange(formulaRange)
    Set formulaContentRange = _
        VTNativeEquationFormulaContentRange(formulaRange)

    If paragraphRange Is Nothing Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & ": native display paragraph is missing."
    End If
    If formulaRange.OMaths(1).Type <> wdOMathDisplay Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & ": OMML is not display math."
    End If
    If sequenceField Is Nothing Or visibleNumberField Is Nothing Or _
       markerRange Is Nothing Or formulaContentRange Is Nothing Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & ": external-SEQ Equation identity is incomplete."
    End If
    If Not VTNativeEquationNumberIsInsideMath( _
       formulaRange, visibleNumberField) Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & ": number REF is outside the OMath array."
    End If
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If paragraphRange.Information(wdWithInTable) Or _
       paragraphRange.Paragraphs.Count <> 1 Or _
       paragraphRange.InlineShapes.Count <> 0 Or _
       paragraphRange.OMaths.Count <> 1 Or _
       paragraphRange.Fields.Count <> 1 Or _
       Not VTHelperParagraphOwnsNativeEquationSequence( _
           helperParagraph) Or _
       helperParagraph.Start < paragraphRange.End Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & _
            ": native Equation formula/SEQ paragraphs are not isolated."
    End If
    For Each candidateField In formulaRange.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            Err.Raise vbObjectError + 7565, "VisualTeX", _
                assertionName & ": native Equation SEQ was absorbed into OMath."
        End If
    Next candidateField
    If paragraphRange.ParagraphFormat.Alignment <> _
           wdAlignParagraphCenter Or _
       VTCustomTabStopCount(paragraphRange) <> 0 Or _
       paragraphRange.ParagraphFormat.KeepWithNext Or _
       paragraphRange.ParagraphFormat.KeepTogether Or _
       paragraphRange.ParagraphFormat.PageBreakBefore Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & _
            ": native Equation paragraph has invalid layout or pagination flags" & _
            " [alignment=" & _
                CStr(paragraphRange.ParagraphFormat.Alignment) & _
            "; customTabs=" & CStr(VTCustomTabStopCount(paragraphRange)) & _
            "; keepWithNext=" & _
                CStr(paragraphRange.ParagraphFormat.KeepWithNext) & _
            "; keepTogether=" & _
                CStr(paragraphRange.ParagraphFormat.KeepTogether) & _
            "; pageBreakBefore=" & _
                CStr(paragraphRange.ParagraphFormat.PageBreakBefore) & "]."
    End If

    fieldStart = VTEquationFieldStart(visibleNumberField)
    fieldEnd = VTEquationFieldEnd(visibleNumberField)
    Set expectedNumberRange = VTNativeEquationArrayNumberRange( _
        formulaRange, visibleNumberField)
    If expectedNumberRange Is Nothing Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & ": native Equation number range is missing."
    End If
    If markerRange.Start >= fieldStart Or _
       fieldEnd > expectedNumberRange.End Or _
       formulaContentRange.End <> markerRange.Start Or _
       StrComp(VTReferenceTargetBookmarkName( _
           visibleNumberField.Code.Text), sequenceBookmarkName, _
           vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & _
            ": native Equation array boundaries or REF target are incomplete."
    End If

    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
    nativeBookmarkName = VTNativeFormulaBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(numberBookmarkName) Or _
       Not documentObject.Bookmarks.Exists(sequenceBookmarkName) Or _
       Not documentObject.Bookmarks.Exists(captionBookmarkName) Or _
       Not documentObject.Bookmarks.Exists(nativeBookmarkName) Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & ": native Equation Bookmark set is incomplete."
    End If
    Set numberRange = documentObject.Bookmarks( _
        numberBookmarkName).Range.Duplicate
    Set sequenceRange = documentObject.Bookmarks( _
        sequenceBookmarkName).Range.Duplicate
    Set captionRange = documentObject.Bookmarks( _
        captionBookmarkName).Range.Duplicate
    Set nativeBookmarkRange = documentObject.Bookmarks( _
        nativeBookmarkName).Range.Duplicate
    expectedNumberText = VTEquationNumberTextForFormula( _
        documentObject, formulaId)
    If Len(expectedNumberText) = 0 Or _
       VTComparableEquationNumberText(sequenceRange.Text) <> _
           VTComparableEquationNumberText(expectedNumberText) Or _
       VTComparableEquationNumberText( _
           VTEquationSequenceResultText(visibleNumberField)) <> _
           VTComparableEquationNumberText(expectedNumberText) Or _
       Not VTNativeEquationNumberBookmarkIsCompatible( _
           numberRange, formulaRange, visibleNumberField) Or _
       nativeBookmarkRange.Start <> formulaRange.Start Or _
       nativeBookmarkRange.End <> formulaRange.End Or _
       Not VTEquationCaptionBookmarkIsCollapsedInParagraph( _
           captionRange, helperParagraph) Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & ": native Equation Bookmark ranges are invalid."
    End If
    If visibleNumberField.Result.Font.Hidden <> False Or _
       visibleNumberField.Result.Font.Color <> wdColorAutomatic Or _
       sequenceField.Result.Font.Hidden <> False Or _
       sequenceField.Result.Font.Color <> wdColorAutomatic Or _
       sequenceField.Result.Font.Position <> 0 Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & ": native Equation REF/SEQ formatting is invalid."
    End If

    nativeItemIndex = VTNativeEquationReferenceItemForFormula( _
        documentObject, formulaId)
    nativeItems = documentObject.GetCrossReferenceItems(wdCaptionEquation)
    If nativeItemIndex < 1 Or Not IsArray(nativeItems) Or _
       nativeItemIndex > UBound(nativeItems) Or _
       VTComparableEquationNumberText( _
           Trim$(CStr(nativeItems(nativeItemIndex)))) <> _
       VTComparableEquationNumberText(expectedNumberText) Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & _
            ": Word's native Equation list does not match VT_N_."
    End If

    documentObject.Repaginate
    textWidth = VTEquationLayoutWidth(paragraphRange)
    If Not VTNativeEquationVisibleHorizontalBounds( _
       formulaContentRange, formulaStartPosition, formulaEndPosition) Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & _
            ": Word did not expose visible native Equation character bounds."
    End If
    Set numberEndProbe = numberRange.Duplicate
    numberEndProbe.Collapse wdCollapseEnd
    numberEndPosition = CSng(numberEndProbe.Information( _
        wdHorizontalPositionRelativeToTextBoundary))
    formulaCenterPosition = _
        (formulaStartPosition + formulaEndPosition) / 2!
    If numberEndPosition < 0! Or _
       Abs(formulaCenterPosition - textWidth / 2!) > 8! Or _
       Abs(numberEndPosition - (textWidth - 1!)) > 4! Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & ": native Equation array geometry is invalid" & _
            " [left=" & CStr(formulaStartPosition) & _
            "; right=" & CStr(formulaEndPosition) & _
            "; center=" & CStr(formulaCenterPosition) & _
            "; numberEnd=" & CStr(numberEndPosition) & _
            "; textWidth=" & CStr(textWidth) & "]."
    End If

    Set formulaProbe = formulaContentRange.Duplicate
    formulaProbe.Collapse wdCollapseStart
    Set numberProbe = visibleNumberField.Result.Duplicate
    numberProbe.Collapse wdCollapseStart
    formulaY = CSng(formulaProbe.Information( _
        wdVerticalPositionRelativeToPage))
    numberY = CSng(numberProbe.Information( _
        wdVerticalPositionRelativeToPage))
    If formulaY < 0! Or numberY < 0! Or _
       Abs(formulaY - numberY) > 6! Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & ": native Equation mathematical axes are not aligned" & _
            " [formulaY=" & CStr(formulaY) & _
            "; numberY=" & CStr(numberY) & "]."
    End If
End Sub

Private Sub VTAssertStaticImageEquationLayout( _
    ByVal formulaRange As Range, _
    ByVal renderedHeightPoints As Double, _
    ByVal formulaId As String, _
    ByVal assertionName As String)

    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim helperParagraph As Range
    Dim sequenceField As Field
    Dim visibleNumberField As Field
    Dim numberRange As Range
    Dim openingRange As Range
    Dim closingRange As Range
    Dim prefixRange As Range
    Dim separatorRange As Range
    Dim suffixRange As Range
    Dim formulaShape As InlineShape
    Dim formulaContainerRange As Range
    Dim textWidth As Single
    Dim numberSize As Single
    Dim numberPosition As Long
    Dim formulaPosition As Long
    Dim expectedPosition As Long
    Dim centerDeltaPoints As Double
    Dim formulaProbe As Range
    Dim numberProbe As Range
    Dim formulaY As Single
    Dim numberY As Single
    Dim formulaCenterY As Single
    Dim numberCenterY As Single
    Dim numberLineHeight As Single
    Dim visualCenterTolerance As Single
    Dim numberFontName As String
    Dim expectedNumber As String

    If formulaRange Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7505, "VisualTeX", _
            assertionName & ": static image formula target is missing."
    End If
    If formulaRange.InlineShapes.Count <> 1 Or _
       formulaRange.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7506, "VisualTeX", _
            assertionName & ": static image formula is ambiguous."
    End If
    Set documentObject = formulaRange.Document
    Set formulaShape = formulaRange.InlineShapes(1)
    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    Set visibleNumberField = VTImageEquationReferenceField( _
        formulaRange, formulaId)
    If Not visibleNumberField Is Nothing Then
        Set numberRange = VTImageEquationNumberRange( _
            formulaRange, visibleNumberField)
    End If
    If sequenceField Is Nothing Or visibleNumberField Is Nothing Or _
       numberRange Is Nothing Then
        Err.Raise vbObjectError + 7506, "VisualTeX", _
            assertionName & ": external SEQ or visible REF range is missing."
    End If
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If helperParagraph.Start <> paragraphRange.End Or _
       Not VTHelperParagraphOwnsNativeEquationSequence(helperParagraph) Then
        Err.Raise vbObjectError + 7506, "VisualTeX", _
            assertionName & ": image SEQ helper is not immediately after the formula row."
    End If
    expectedNumber = VTEquationNumberTextForFormula( _
        documentObject, formulaId)
    Set prefixRange = documentObject.Range( _
        Start:=paragraphRange.Start, End:=formulaContainerRange.Start)
    Set separatorRange = documentObject.Range( _
        Start:=formulaContainerRange.End, End:=numberRange.Start)
    Set suffixRange = documentObject.Range( _
        Start:=numberRange.End, End:=paragraphRange.End)
    If Not VTParagraphHasSingleVisualTeXImageMacroButton( _
           paragraphRange) Or _
       prefixRange.Text <> vbTab Or _
       separatorRange.Text <> vbTab Or _
       suffixRange.Text <> vbCr Or _
       numberRange.Text <> "(" & expectedNumber & ")" Then
        Err.Raise vbObjectError + 7512, "VisualTeX", _
            assertionName & _
            ": static image number paragraph contains invalid content."
    End If

    textWidth = VTEquationLayoutWidth(paragraphRange)
    If paragraphRange.ParagraphFormat.Alignment <> wdAlignParagraphJustify Or _
       paragraphRange.ParagraphFormat.TabStops.Count <> 2 Or _
       paragraphRange.ParagraphFormat.TabStops(1).Alignment <> _
           wdAlignTabCenter Or _
       paragraphRange.ParagraphFormat.TabStops(2).Alignment <> _
           wdAlignTabRight Or _
       Abs(paragraphRange.ParagraphFormat.TabStops(1).Position - _
           textWidth / 2!) > 0.5 Or _
       Abs(paragraphRange.ParagraphFormat.TabStops(2).Position - _
           (textWidth - 1!)) > 0.5 Then
        Err.Raise vbObjectError + 7509, "VisualTeX", _
            assertionName & ": static image tab geometry is invalid."
    End If
    If renderedHeightPoints > 0! And _
       Abs(formulaShape.Height - renderedHeightPoints) > 0.5 Then
        Err.Raise vbObjectError + 7514, "VisualTeX", _
            assertionName & ": image height changed unexpectedly."
    End If
    Set openingRange = documentObject.Range( _
        Start:=numberRange.Start, End:=numberRange.Start + 1)
    Set closingRange = documentObject.Range( _
        Start:=numberRange.End - 1, End:=numberRange.End)
    numberSize = visibleNumberField.Result.Font.Size
    numberPosition = numberRange.Font.Position
    formulaPosition = formulaShape.Range.Font.Position
    numberFontName = CStr(visibleNumberField.Result.Font.NameAscii)
    If numberSize <= 0! Or _
       numberSize > CSng(VT_WORD_MAX_FORMULA_FONT_SIZE_PT) Or _
       numberPosition <> 0 Or _
       formulaPosition = wdUndefined Or formulaPosition < -256 Or _
       formulaPosition > 256 Or _
       StrComp(numberFontName, VT_WORD_EQUATION_NUMBER_FONT_NAME, _
           vbTextCompare) <> 0 Or _
       StrComp(CStr(visibleNumberField.Result.Font.NameOther), _
           VT_WORD_EQUATION_NUMBER_FONT_NAME, vbTextCompare) <> 0 Or _
       StrComp(CStr(openingRange.Font.NameAscii), _
           VT_WORD_EQUATION_NUMBER_FONT_NAME, vbTextCompare) <> 0 Or _
       StrComp(CStr(closingRange.Font.NameAscii), _
           VT_WORD_EQUATION_NUMBER_FONT_NAME, vbTextCompare) <> 0 Or _
       visibleNumberField.Result.Font.Hidden <> False Or _
       visibleNumberField.Result.Font.Color <> wdColorAutomatic Then
        Err.Raise vbObjectError + 7513, "VisualTeX", _
            assertionName & _
            ": image object/visible-REF formatting is invalid" & _
            " [formulaPosition=" & CStr(formulaPosition) & _
            "; numberPosition=" & CStr(numberPosition) & _
            "; resultAscii=" & numberFontName & _
            "; resultOther=" & CStr( _
                visibleNumberField.Result.Font.NameOther) & _
            "; openAscii=" & CStr(openingRange.Font.NameAscii) & _
            "; closeAscii=" & CStr(closingRange.Font.NameAscii) & "]."
    End If

    documentObject.Repaginate
    If Not VTTryCalculateCurrentNumberedImagePaintedCenterDeltaPoints( _
       formulaShape, numberRange, centerDeltaPoints) Then
        Err.Raise vbObjectError + 7521, "VisualTeX", _
            assertionName & ": Word did not expose final ink-center geometry."
    End If
    If Abs(centerDeltaPoints) > 0.5001# Then
        Err.Raise vbObjectError + 7521, "VisualTeX", _
            assertionName & ": image and number ink centers exceed Position quantisation" & _
            " [deltaPoints=" & CStr(centerDeltaPoints) & _
            "; formulaPosition=" & CStr(formulaPosition) & "]."
    End If
End Sub

Private Sub VTAssertNumberedEquationLayout( _
    ByVal formulaRange As Range, _
    ByVal renderedHeightPoints As Double, _
    ByVal formulaId As String, _
    ByVal expectedCaptionText As String, _
    ByVal assertionName As String)

    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim helperParagraph As Range
    Dim sequenceField As Field
    Dim visibleNumberField As Field
    Dim openingRange As Range
    Dim resultRange As Range
    Dim closingRange As Range
    Dim numberRange As Range
    Dim sequenceRange As Range
    Dim captionRange As Range
    Dim prefixRange As Range
    Dim separatorRange As Range
    Dim suffixRange As Range
    Dim formulaStartProbe As Range
    Dim formulaEndProbe As Range
    Dim numberEndProbe As Range
    Dim formulaProbe As Range
    Dim numberProbe As Range
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim textWidth As Single
    Dim preferredSize As Single
    Dim expectedRaise As Long
    Dim formulaStartPosition As Single
    Dim formulaEndPosition As Single
    Dim formulaCenterPosition As Single
    Dim numberEndPosition As Single
    Dim centerTolerance As Single
    Dim nativeFontSize As Single
    Dim formulaLine As Long
    Dim numberLine As Long
    Dim formulaY As Single
    Dim numberY As Single
    Dim formulaCenterY As Single
    Dim numberCenterY As Single
    Dim numberLineHeight As Single
    Dim visualCenterTolerance As Single
    Dim numberPosition As Long
    Dim numberBookmarkName As String
    Dim sequenceBookmarkName As String
    Dim captionBookmarkName As String

    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7505, "VisualTeX", _
            assertionName & ": formula Range is missing."
    End If
    If formulaRange.Information(wdWithInTable) Then
        VTAssertNumberedDisplayTableLayout _
            formulaRange, renderedHeightPoints, formulaId, _
            expectedCaptionText, assertionName
        Exit Sub
    End If
    If formulaRange.OMaths.Count = 1 Then
        VTAssertNativeEquationArrayLayout _
            formulaRange, formulaId, expectedCaptionText, assertionName
        Exit Sub
    End If
    If formulaRange.InlineShapes.Count = 1 Then
        VTAssertStaticImageEquationLayout _
            formulaRange, renderedHeightPoints, formulaId, assertionName
        Exit Sub
    End If

    Set documentObject = formulaRange.Document
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    If paragraphRange.Information(wdWithInTable) Or _
       paragraphRange.Paragraphs.Count <> 1 Or _
       paragraphRange.InlineShapes.Count + paragraphRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7506, "VisualTeX", _
            assertionName & ": numbered formula is not one ordinary paragraph."
    End If
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    Set visibleNumberField = VTImageEquationReferenceField( _
        formulaRange, formulaId)
    If sequenceField Is Nothing Or visibleNumberField Is Nothing Then
        Err.Raise vbObjectError + 7506, "VisualTeX", _
            assertionName & ": image Equation SEQ/REF identity is missing."
    End If
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If Not VTHelperParagraphOwnsNativeEquationSequence( _
           helperParagraph) Or _
       helperParagraph.Start < paragraphRange.End Then
        Err.Raise vbObjectError + 7506, "VisualTeX", _
            assertionName & ": image Equation SEQ helper is invalid."
    End If

    textWidth = VTEquationLayoutWidth(paragraphRange)
    If paragraphRange.ParagraphFormat.Alignment <> wdAlignParagraphLeft Or _
       paragraphRange.ParagraphFormat.KeepWithNext Or _
       paragraphRange.ParagraphFormat.KeepTogether Or _
       paragraphRange.ParagraphFormat.PageBreakBefore Then
        Err.Raise vbObjectError + 7507, "VisualTeX", _
            assertionName & _
            ": numbered image paragraph has invalid alignment or pagination flags."
    End If
    If paragraphRange.ParagraphFormat.TabStops.Count <> 2 Then
        Err.Raise vbObjectError + 7508, "VisualTeX", _
            assertionName & ": numbered paragraph does not have exactly two tab stops."
    End If
    If paragraphRange.ParagraphFormat.TabStops(1).Alignment <> _
       wdAlignTabCenter Or _
       Abs(paragraphRange.ParagraphFormat.TabStops(1).Position - _
           textWidth / 2!) > 0.5 Then
        Err.Raise vbObjectError + 7509, "VisualTeX", _
            assertionName & ": formula center tab is not at the text-column center."
    End If
    If paragraphRange.ParagraphFormat.TabStops(2).Alignment <> _
       wdAlignTabRight Or _
       Abs(paragraphRange.ParagraphFormat.TabStops(2).Position - _
           (textWidth - 1!)) > 0.5 Then
        Err.Raise vbObjectError + 7510, "VisualTeX", _
            assertionName & ": Equation number tab is not at the right text boundary."
    End If
    If documentObject.Range( _
        paragraphRange.Start, paragraphRange.Start + 1).Text <> vbTab Then
        Err.Raise vbObjectError + 7511, "VisualTeX", _
            assertionName & ": formula is not anchored by the center tab."
    End If

    fieldStart = VTEquationFieldStart(visibleNumberField)
    fieldEnd = VTEquationFieldEnd(visibleNumberField)
    If fieldStart <= formulaRange.End Or _
       fieldEnd >= paragraphRange.End Or _
       fieldStart < 2 Or _
       documentObject.Range(fieldStart - 2, fieldStart - 1).Text <> vbTab Then
        Err.Raise vbObjectError + 7512, "VisualTeX", _
            assertionName & _
            ": Equation number is not outside the formula at the right tab."
    End If
    Set prefixRange = documentObject.Range( _
        Start:=paragraphRange.Start, End:=formulaRange.Start)
    Set separatorRange = documentObject.Range( _
        Start:=formulaRange.End, End:=fieldStart)
    Set suffixRange = documentObject.Range( _
        Start:=fieldEnd, End:=paragraphRange.End)
    If paragraphRange.Fields.Count <> 1 Or _
       prefixRange.Text <> vbTab Or _
       separatorRange.Text <> vbTab & "(" Or _
       suffixRange.Text <> ")" & vbCr Then
        Err.Raise vbObjectError + 7512, "VisualTeX", _
            assertionName & _
            ": numbered paragraph contains extra visible or hidden content" & _
            " [fields=" & CStr(paragraphRange.Fields.Count) & _
            "; prefix=" & Replace$(prefixRange.Text, vbTab, "<TAB>") & _
            "; separator=" & Replace$(separatorRange.Text, vbTab, "<TAB>") & _
            "; suffix=" & Replace$(suffixRange.Text, vbCr, "<CR>") & "]."
    End If
    Set openingRange = documentObject.Range( _
        Start:=fieldStart - 1, End:=fieldStart)
    Set closingRange = documentObject.Range( _
        Start:=fieldEnd, End:=fieldEnd + 1)
    If openingRange.Text <> "(" Or closingRange.Text <> ")" Then
        Err.Raise vbObjectError + 7513, "VisualTeX", _
            assertionName & ": Equation number parentheses are incomplete."
    End If

    Set resultRange = visibleNumberField.Result.Duplicate
    preferredSize = resultRange.Font.Size
    If preferredSize <= 0! Or preferredSize > 72! Or _
       openingRange.Font.Size <> preferredSize Or _
       closingRange.Font.Size <> preferredSize Or _
       openingRange.Font.Hidden <> False Or _
       resultRange.Font.Hidden <> False Or _
       closingRange.Font.Hidden <> False Or _
       openingRange.Font.Color <> wdColorAutomatic Or _
       resultRange.Font.Color <> wdColorAutomatic Or _
       closingRange.Font.Color <> wdColorAutomatic Then
        Err.Raise vbObjectError + 7514, "VisualTeX", _
            assertionName & ": Equation number is not the only normally visible text."
    End If
    expectedRaise = openingRange.Font.Position
    If openingRange.Font.Position <> resultRange.Font.Position Or _
       openingRange.Font.Position <> closingRange.Font.Position Or _
       expectedRaise = wdUndefined Or expectedRaise < -48 Or _
       expectedRaise > 48 Then
        Err.Raise vbObjectError + 7515, "VisualTeX", _
            assertionName & ": Equation number is not vertically stable" & _
            " [opening=" & CStr(openingRange.Font.Position) & _
            "; result=" & CStr(resultRange.Font.Position) & _
            "; closing=" & CStr(closingRange.Font.Position) & _
            "; expected=" & CStr(expectedRaise) & _
            "; image=" & CStr(formulaRange.InlineShapes.Count = 1) & _
            "; omml=" & CStr(formulaRange.OMaths.Count = 1) & "]."
    End If

    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    sequenceBookmarkName = VTEquationSequenceNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(numberBookmarkName) Or _
       Not documentObject.Bookmarks.Exists(sequenceBookmarkName) Or _
       Not documentObject.Bookmarks.Exists(captionBookmarkName) Then
        Err.Raise vbObjectError + 7516, "VisualTeX", _
            assertionName & ": Equation Bookmark set is incomplete."
    End If
    Set numberRange = documentObject.Bookmarks( _
        numberBookmarkName).Range.Duplicate
    Set sequenceRange = documentObject.Bookmarks( _
        sequenceBookmarkName).Range.Duplicate
    Set captionRange = documentObject.Bookmarks( _
        captionBookmarkName).Range.Duplicate
    If numberRange.Text <> "(" & resultRange.Text & ")" Or _
       sequenceRange.Text <> resultRange.Text Or _
       numberRange.Paragraphs(1).Range.Start <> paragraphRange.Start Or _
       sequenceRange.Paragraphs(1).Range.Start <> helperParagraph.Start Or _
       captionRange.Start <> captionRange.End Or _
       captionRange.Paragraphs(1).Range.Start <> helperParagraph.Start Or _
       InStr(1, numberRange.Text, expectedCaptionText, vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 7517, "VisualTeX", _
            assertionName & ": Equation Bookmark ranges are invalid" & _
            " [number=" & Replace$(numberRange.Text, vbCr, "<CR>") & _
            "; sequence=" & Replace$(sequenceRange.Text, vbCr, "<CR>") & _
            "; caption=" & CStr(captionRange.Start) & "-" & _
                CStr(captionRange.End) & "]."
    End If

    documentObject.Repaginate
    Set formulaStartProbe = formulaRange.Duplicate
    formulaStartProbe.Collapse wdCollapseStart
    Set formulaEndProbe = formulaRange.Duplicate
    formulaEndProbe.Collapse wdCollapseEnd
    Set numberEndProbe = closingRange.Duplicate
    numberEndProbe.Collapse wdCollapseEnd
    formulaStartPosition = CSng(formulaStartProbe.Information( _
        wdHorizontalPositionRelativeToTextBoundary))
    formulaEndPosition = CSng(formulaEndProbe.Information( _
        wdHorizontalPositionRelativeToTextBoundary))
    numberEndPosition = CSng(numberEndProbe.Information( _
        wdHorizontalPositionRelativeToTextBoundary))
    If formulaStartPosition < 0! Or formulaEndPosition < 0! Or _
       numberEndPosition < 0! Then
        Err.Raise vbObjectError + 7518, "VisualTeX", _
            assertionName & ": Word did not expose measurable layout positions."
    End If
    formulaCenterPosition = _
        (formulaStartPosition + formulaEndPosition) / 2!
    centerTolerance = 3!
    If formulaRange.OMaths.Count = 1 Then
        nativeFontSize = formulaRange.Font.Size
        If nativeFontSize > 0! And nativeFontSize <= 72! Then
            If centerTolerance < nativeFontSize / 2! Then
                centerTolerance = nativeFontSize / 2!
            End If
        End If
    End If
    If Abs(formulaCenterPosition - textWidth / 2!) > centerTolerance Then
        Err.Raise vbObjectError + 7519, "VisualTeX", _
            assertionName & ": formula is not geometrically centered" & _
            " [center=" & CStr(formulaCenterPosition) & _
            "; target=" & CStr(textWidth / 2!) & _
            "; tolerance=" & CStr(centerTolerance) & "]."
    End If
    If Abs(numberEndPosition - (textWidth - 1!)) > 3! Then
        Err.Raise vbObjectError + 7520, "VisualTeX", _
            assertionName & ": Equation number is not at the right text boundary" & _
            " [numberEnd=" & CStr(numberEndPosition) & _
            "; target=" & CStr(textWidth - 1!) & "]."
    End If
    formulaLine = formulaStartProbe.Information(wdFirstCharacterLineNumber)
    numberLine = numberEndProbe.Information(wdFirstCharacterLineNumber)
    If formulaLine <= 0 Or numberLine <= 0 Or formulaLine <> numberLine Then
        Err.Raise vbObjectError + 7521, "VisualTeX", _
            assertionName & ": formula and Equation number are not on the same line."
    End If

    Set formulaProbe = formulaRange.Duplicate
    formulaProbe.Collapse wdCollapseStart
    Set numberProbe = numberRange.Duplicate
    numberProbe.Collapse wdCollapseStart
    formulaY = CSng(formulaProbe.Information( _
        wdVerticalPositionRelativeToPage))
    numberY = CSng(numberProbe.Information( _
        wdVerticalPositionRelativeToPage))
    If formulaY < 0! Or numberY < 0! Then
        Err.Raise vbObjectError + 7521, "VisualTeX", _
            assertionName & ": Word did not expose vertical layout positions."
    End If
    numberLineHeight = numberRange.ParagraphFormat.LineSpacing
    If numberLineHeight <= 0! Or numberLineHeight = wdUndefined Or _
       numberLineHeight > 72! Then
        numberLineHeight = numberRange.Font.Size * 1.2!
    End If
    If numberLineHeight <= 0! Or numberLineHeight > 72! Then
        numberLineHeight = 14.4!
    End If
    numberPosition = numberRange.Font.Position
    If numberPosition = wdUndefined Then numberPosition = 0
    formulaCenterY = formulaY - CSng(renderedHeightPoints / 2#)
    numberCenterY = numberY - numberLineHeight / 2! - numberPosition
    visualCenterTolerance = 2!
    If Abs(formulaCenterY - numberCenterY) > visualCenterTolerance Then
        Err.Raise vbObjectError + 7521, "VisualTeX", _
            assertionName & _
            ": image formula and number visual centers are not aligned" & _
            " [formulaBaseline=" & CStr(formulaY) & _
            "; formulaHeight=" & CStr(renderedHeightPoints) & _
            "; formulaCenter=" & CStr(formulaCenterY) & _
            "; numberBaseline=" & CStr(numberY) & _
            "; numberLineHeight=" & CStr(numberLineHeight) & _
            "; numberPosition=" & CStr(numberPosition) & _
            "; numberCenter=" & CStr(numberCenterY) & _
            "; tolerance=" & CStr(visualCenterTolerance) & "]."
    End If
End Sub

Private Sub VTAssertNumberedDisplayTableLayout( _
    ByVal formulaRange As Range, _
    ByVal renderedHeightPoints As Double, _
    ByVal formulaId As String, _
    ByVal expectedCaptionText As String, _
    ByVal assertionName As String)

    Dim documentObject As Document
    Dim layoutTable As Table
    Dim sequenceField As Field
    Dim referenceField As Field
    Dim candidateField As Field
    Dim numberRange As Range
    Dim mirrorRange As Range
    Dim sequenceRange As Range
    Dim trailingParagraph As Range
    Dim formulaProbe As Range
    Dim numberProbe As Range
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim equationLabelName As String
    Dim expectedNumberText As String
    Dim visibleText As String
    Dim cellXml As String
    Dim normalSize As Single
    Dim formulaY As Single
    Dim numberY As Single
    Dim formulaCenterY As Single
    Dim numberCenterY As Single
    Dim numberLineHeight As Single
    Dim visualCenterTolerance As Single
    Dim cellIndex As Long
    Dim sequenceOrdinal As Long

    Set documentObject = formulaRange.Document
    Set layoutTable = formulaRange.Tables(1)
    If layoutTable.Rows.Count <> 1 Or layoutTable.Columns.Count <> 3 Or _
       layoutTable.AllowAutoFit Or _
       layoutTable.PreferredWidthType <> wdPreferredWidthPercent Or _
       (layoutTable.PreferredWidth <> wdUndefined And _
        Abs(layoutTable.PreferredWidth - 100!) > 0.1) Or _
       layoutTable.Borders.Enable <> 0 Then
        Err.Raise vbObjectError + 7507, "VisualTeX", _
            assertionName & ": numbered formula table geometry is invalid" & _
            " [rows=" & CStr(layoutTable.Rows.Count) & _
            "; columns=" & CStr(layoutTable.Columns.Count) & _
            "; autoFit=" & CStr(layoutTable.AllowAutoFit) & _
            "; widthType=" & CStr(layoutTable.PreferredWidthType) & _
            "; width=" & CStr(layoutTable.PreferredWidth) & _
            "; borders=" & CStr(layoutTable.Borders.Enable) & _
            "; allowBreak=" & _
                CStr(layoutTable.Rows.AllowBreakAcrossPages) & "]."
    End If
    For cellIndex = 1 To 3
        If layoutTable.Cell(1, cellIndex).VerticalAlignment <> _
           wdCellAlignVerticalCenter Then
            Err.Raise vbObjectError + 7508, "VisualTeX", _
                assertionName & ": a numbered formula cell is not vertically centered."
        End If
    Next cellIndex
    If Abs(layoutTable.Columns(1).PreferredWidth - 20!) > 0.1 Or _
       Abs(layoutTable.Columns(2).PreferredWidth - 60!) > 0.1 Or _
       Abs(layoutTable.Columns(3).PreferredWidth - 20!) > 0.1 Or _
       layoutTable.Cell(1, 2).Range.ParagraphFormat.Alignment <> _
           wdAlignParagraphCenter Or _
       layoutTable.Cell(1, 3).Range.ParagraphFormat.Alignment <> _
           wdAlignParagraphRight Then
        Err.Raise vbObjectError + 7509, "VisualTeX", _
            assertionName & ": numbered formula table is not 20/60/20 centered layout."
    End If
    If formulaRange.Start < layoutTable.Cell(1, 2).Range.Start Or _
       formulaRange.End > layoutTable.Cell(1, 2).Range.End Then
        Err.Raise vbObjectError + 7510, "VisualTeX", _
            assertionName & ": formula is not in the center table cell."
    End If

    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(sequenceBookmarkName) Or _
       Not documentObject.Bookmarks.Exists(numberBookmarkName) Then
        Err.Raise vbObjectError + 7516, "VisualTeX", _
            assertionName & ": hidden SEQ or complete number mirror is missing."
    End If
    Set sequenceRange = documentObject.Bookmarks( _
        sequenceBookmarkName).Range.Duplicate
    Set mirrorRange = documentObject.Bookmarks( _
        numberBookmarkName).Range.Duplicate
    For Each candidateField In documentObject.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            If candidateField.Result.Start <= sequenceRange.Start And _
               candidateField.Result.End >= sequenceRange.End Then
                Set sequenceField = candidateField
                Exit For
            End If
        End If
    Next candidateField
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7506, "VisualTeX", _
            assertionName & ": hidden native Equation SEQ field is missing."
    End If
    equationLabelName = VTNativeEquationLabelName()
    sequenceOrdinal = VTEquationSequenceOrdinal( _
        documentObject, sequenceField, equationLabelName)
    If sequenceOrdinal < 1 Or _
       Not VTEquationSequenceFieldHasOrdinal( _
           sequenceField, equationLabelName, sequenceOrdinal) Then
        Err.Raise vbObjectError + 7517, "VisualTeX", _
            assertionName & ": hidden native Equation SEQ code is invalid" & _
            " [code=" & sequenceField.Code.Text & _
            "; result=" & sequenceField.Result.Text & "]."
    End If
    expectedNumberText = "(" & CStr(sequenceOrdinal) & ")"
    For Each candidateField In layoutTable.Cell(1, 3).Range.Fields
        If candidateField.Type = wdFieldRef And _
           InStr(1, candidateField.Code.Text, sequenceBookmarkName, _
           vbTextCompare) > 0 Then
            Set referenceField = candidateField
            Exit For
        End If
    Next candidateField
    If referenceField Is Nothing Then
        Err.Raise vbObjectError + 7516, "VisualTeX", _
            assertionName & ": right cell native REF field is missing."
    End If
    Set numberRange = mirrorRange.Duplicate
    visibleText = Trim$(numberRange.Text)
    normalSize = VTVisibleEquationNumberFontSize(documentObject)
    If visibleText <> expectedNumberText Or _
       Trim$(referenceField.Result.Text) <> CStr(sequenceOrdinal) Or _
       InStr(1, visibleText, expectedCaptionText, vbTextCompare) > 0 Or _
       InStr(1, referenceField.Code.Text, "MERGEFORMAT", _
           vbTextCompare) > 0 Or _
       referenceField.Result.Font.Hidden <> False Or _
       referenceField.Result.Font.Color <> wdColorAutomatic Or _
       Abs(referenceField.Result.Font.Size - normalSize) > 0.1 Or _
       numberRange.Font.Hidden <> False Or _
       numberRange.Font.Color <> wdColorAutomatic Or _
       numberRange.Font.Position <> 0 Then
        Err.Raise vbObjectError + 7517, "VisualTeX", _
            assertionName & ": visible Equation number is not truly visible" & _
            " [text=" & visibleText & _
            "; code=" & referenceField.Code.Text & _
            "; size=" & CStr(referenceField.Result.Font.Size) & _
            "; expectedSize=" & CStr(normalSize) & _
            "; color=" & CStr(referenceField.Result.Font.Color) & _
            "; hidden=" & CStr(referenceField.Result.Font.Hidden) & _
            "; position=" & CStr(numberRange.Font.Position) & "]."
    End If
    If sequenceRange.Text <> sequenceField.Result.Text Or _
       InStr(1, sequenceRange.Text, expectedCaptionText, _
       vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 7517, "VisualTeX", _
            assertionName & ": hidden Equation sequence contains formula text."
    End If

    If formulaRange.OMaths.Count = 1 Then
        On Error Resume Next
        normalSize = documentObject.Styles(wdStyleNormal).Font.Size
        On Error GoTo 0
        If normalSize <= 0! Or normalSize > 72! Then normalSize = 12!
        If formulaRange.OMaths(1).Type <> wdOMathDisplay Or _
           formulaRange.Font.Position <> 0 Or _
           Abs(formulaRange.Font.Size - normalSize) > 1.5 Then
            Err.Raise vbObjectError + 7517, "VisualTeX", _
                assertionName & ": OMML is not authentic Word display math" & _
                " [type=" & CStr(formulaRange.OMaths(1).Type) & _
                "; size=" & CStr(formulaRange.Font.Size) & _
                "; normal=" & CStr(normalSize) & _
                "; position=" & CStr(formulaRange.Font.Position) & "]."
        End If
        cellXml = VTProbeRangeWordOpenXml( _
            layoutTable.Cell(1, 2).Range)
        If layoutTable.Cell(1, 2).Range.Paragraphs.Count <> 2 Or _
           InStr(1, cellXml, "<m:oMathPara", vbBinaryCompare) = 0 Then
            Err.Raise vbObjectError + 7517, "VisualTeX", _
                assertionName & _
                ": numbered OMML lost its two-paragraph m:oMathPara cell."
        End If
        Set trailingParagraph = _
            layoutTable.Cell(1, 2).Range.Paragraphs(2).Range.Duplicate
        If trailingParagraph.OMaths.Count <> 0 Or _
           VTWordRangeHasMeaningfulText(trailingParagraph) Or _
           trailingParagraph.Font.Size <> 1! Or _
           trailingParagraph.ParagraphFormat.LineSpacingRule <> _
               wdLineSpaceExactly Or _
           Abs(trailingParagraph.ParagraphFormat.LineSpacing - 1!) > 0.1 Then
            Err.Raise vbObjectError + 7517, "VisualTeX", _
                assertionName & _
                ": numbered OMML display tail is not the compact 1pt paragraph."
        End If
    End If

    documentObject.Repaginate
    Set formulaProbe = formulaRange.Duplicate
    formulaProbe.Collapse wdCollapseStart
    Set numberProbe = numberRange.Duplicate
    numberProbe.Collapse wdCollapseStart
    formulaY = CSng(formulaProbe.Information( _
        wdVerticalPositionRelativeToPage))
    numberY = CSng(numberProbe.Information( _
        wdVerticalPositionRelativeToPage))
    If formulaY < 0! Or numberY < 0! Then
        Err.Raise vbObjectError + 7521, "VisualTeX", _
            assertionName & ": Word did not expose vertical layout positions" & _
            " [formulaY=" & CStr(formulaY) & _
            "; numberY=" & CStr(numberY) & "]."
    End If

    If formulaRange.InlineShapes.Count = 1 Then
        ' A collapsed InlineShape Range reports the top of the image line box,
        ' while the number Range reports the top of its text line. Comparing
        ' those origins falsely rejects tall formulas. Compare their visual
        ' centers using the rendered image height and Word's actual number-line
        ' spacing instead; this mirrors the validated Windows screen-rectangle
        ' center check without relying on a Windows-only GetPoint call.
        On Error Resume Next
        numberLineHeight = numberRange.ParagraphFormat.LineSpacing
        On Error GoTo 0
        If numberLineHeight <= 0! Or numberLineHeight = wdUndefined Or _
           numberLineHeight > 72! Then
            numberLineHeight = numberRange.Font.Size * 1.2!
        End If
        If numberLineHeight <= 0! Or numberLineHeight > 72! Then
            numberLineHeight = 14.4!
        End If
        formulaCenterY = formulaY + CSng(renderedHeightPoints / 2#)
        numberCenterY = numberY + numberLineHeight / 2!
        visualCenterTolerance = 4!
        If Abs(formulaCenterY - numberCenterY) > visualCenterTolerance Then
            Err.Raise vbObjectError + 7521, "VisualTeX", _
                assertionName & _
                ": formula and number visual centers are not aligned" & _
                " [formulaY=" & CStr(formulaY) & _
                "; formulaHeight=" & CStr(renderedHeightPoints) & _
                "; formulaCenter=" & CStr(formulaCenterY) & _
                "; numberY=" & CStr(numberY) & _
                "; numberLineHeight=" & CStr(numberLineHeight) & _
                "; numberCenter=" & CStr(numberCenterY) & _
                "; tolerance=" & CStr(visualCenterTolerance) & "]."
        End If
    ElseIf Abs(formulaY - numberY) > 18! Then
        Err.Raise vbObjectError + 7521, "VisualTeX", _
            assertionName & ": formula and number do not share a visual axis" & _
            " [formulaY=" & CStr(formulaY) & _
            "; numberY=" & CStr(numberY) & "]."
    End If
End Sub

Private Function VTNumberedEquationInvariantSnapshot( _
    ByVal documentObject As Document, _
    ByVal paragraphStart As Long) As String

    Dim paragraphRange As Range
    Dim formulaRange As Range
    Dim helperParagraph As Range
    Dim numberRange As Range
    Dim sequenceField As Field
    Dim visibleNumberField As Field
    Dim openingRange As Range
    Dim closingRange As Range
    Dim formulaStartProbe As Range
    Dim formulaEndProbe As Range
    Dim numberEndProbe As Range
    Dim numberBookmark As Bookmark
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim formulaType As Long
    Dim bookmarkName As String
    Dim formulaId As String
    Dim sequenceBookmarkName As String

    If documentObject Is Nothing Or paragraphStart < 0 Or _
       paragraphStart >= documentObject.Content.End Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The numbered Equation invariant target is invalid."
    End If
    Set paragraphRange = documentObject.Range( _
        Start:=paragraphStart, End:=paragraphStart).Paragraphs(1).Range.Duplicate
    If paragraphRange.Information(wdWithInTable) Then
        VTNumberedEquationInvariantSnapshot = _
            VTNumberedDisplayTableInvariantSnapshot(paragraphRange.Tables(1))
        Exit Function
    End If
    If paragraphRange.OMaths.Count = 1 Then
        Set formulaRange = VTNativeMathRoot(paragraphRange.OMaths(1)).Range.Duplicate
        formulaType = paragraphRange.OMaths(1).Type
        bookmarkName = ""
        For Each numberBookmark In documentObject.Bookmarks
            If Left$(numberBookmark.Name, _
               Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
               VT_WORD_NUMBER_BOOKMARK_PREFIX And _
               Not numberBookmark.Range.Information(wdWithInTable) Then
                If numberBookmark.Range.Paragraphs(1).Range.Start = _
                   paragraphRange.Start Then
                    formulaId = VTFormulaIdFromBookmarkSuffix( _
                        Mid$(numberBookmark.Name, _
                            Len(VT_WORD_NUMBER_BOOKMARK_PREFIX) + 1))
                    If Len(formulaId) > 0 Then
                        Set visibleNumberField = _
                            VTNativeEquationArrayReferenceField( _
                                formulaRange, formulaId)
                        If Not visibleNumberField Is Nothing Then
                            If VTNativeEquationNumberBookmarkIsCompatible( _
                               numberBookmark.Range.Duplicate, formulaRange, _
                               visibleNumberField) Then
                                bookmarkName = numberBookmark.Name
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next numberBookmark
        If Len(bookmarkName) = 0 Or visibleNumberField Is Nothing Then
            Err.Raise vbObjectError + 7547, "VisualTeX", _
                "The numbered native Equation invariant REF is missing."
        End If
        sequenceBookmarkName = _
            VTEquationSequenceNumberBookmarkName(formulaId)
        Set sequenceField = VTEquationSequenceFieldForBookmark( _
            documentObject, sequenceBookmarkName)
        If sequenceField Is Nothing Then
            Err.Raise vbObjectError + 7547, "VisualTeX", _
                "The numbered native Equation invariant SEQ is missing."
        End If
        Set helperParagraph = _
            sequenceField.Result.Paragraphs(1).Range.Duplicate
        If Not VTHelperParagraphOwnsNativeEquationSequence( _
           helperParagraph) Or helperParagraph.Start < paragraphRange.End Then
            Err.Raise vbObjectError + 7547, "VisualTeX", _
                "The numbered native Equation invariant helper is invalid."
        End If
        Set numberRange = documentObject.Bookmarks( _
            bookmarkName).Range.Duplicate
        Set formulaStartProbe = formulaRange.Duplicate
        formulaStartProbe.Collapse wdCollapseStart
        Set formulaEndProbe = formulaRange.Duplicate
        formulaEndProbe.Collapse wdCollapseEnd
        Set numberEndProbe = numberRange.Duplicate
        numberEndProbe.Collapse wdCollapseEnd
        documentObject.Repaginate
        VTNumberedEquationInvariantSnapshot = _
            "paragraph=" & CStr(paragraphRange.Start) & ":" & _
                CStr(paragraphRange.End) & _
            "|formula=" & CStr(formulaRange.Start) & ":" & _
                CStr(formulaRange.End) & ":" & CStr(formulaType) & _
                ":" & CStr(formulaRange.Font.Size) & _
            "|visibleRef=" & CStr(VTEquationFieldStart( _
                visibleNumberField)) & ":" & _
                CStr(VTEquationFieldEnd(visibleNumberField)) & ":" & _
                Trim$(visibleNumberField.Code.Text) & ":" & _
                visibleNumberField.Result.Text
        VTNumberedEquationInvariantSnapshot = _
            VTNumberedEquationInvariantSnapshot & _
            "|externalSeq=" & CStr(VTEquationFieldStart(sequenceField)) & _
                ":" & CStr(VTEquationFieldEnd(sequenceField)) & ":" & _
                Trim$(sequenceField.Code.Text) & ":" & _
                sequenceField.Result.Text & _
            "|helper=" & CStr(helperParagraph.Start) & ":" & _
                CStr(helperParagraph.End) & _
            "|bookmark=" & bookmarkName & ":" & numberRange.Text & _
            "|tabs=" & CStr(VTCustomTabStopCount(paragraphRange))
        VTNumberedEquationInvariantSnapshot = _
            VTNumberedEquationInvariantSnapshot & _
            "|xy=" & CStr(formulaStartProbe.Information( _
                wdHorizontalPositionRelativeToTextBoundary)) & ":" & _
                CStr(formulaEndProbe.Information( _
                wdHorizontalPositionRelativeToTextBoundary)) & ":" & _
                CStr(numberEndProbe.Information( _
                wdHorizontalPositionRelativeToTextBoundary)) & _
            "|line=" & CStr(formulaStartProbe.Information( _
                wdFirstCharacterLineNumber)) & ":" & _
                CStr(numberEndProbe.Information(wdFirstCharacterLineNumber))
        Exit Function
    End If

    If paragraphRange.InlineShapes.Count = 1 Then
        Set formulaRange = paragraphRange.InlineShapes(1).Range.Duplicate
        formulaType = -1
    Else
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The numbered Equation invariant formula is ambiguous."
    End If

    bookmarkName = ""
    For Each numberBookmark In documentObject.Bookmarks
        If Left$(numberBookmark.Name, _
           Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_NUMBER_BOOKMARK_PREFIX And _
           Not numberBookmark.Range.Information(wdWithInTable) Then
            If numberBookmark.Range.Paragraphs(1).Range.Start = _
               paragraphRange.Start Then
                formulaId = VTFormulaIdFromBookmarkSuffix( _
                    Mid$(numberBookmark.Name, _
                        Len(VT_WORD_NUMBER_BOOKMARK_PREFIX) + 1))
                If Len(formulaId) > 0 Then
                    Set numberRange = VTStaticImageEquationNumberRange( _
                        formulaRange, formulaId)
                    If Not numberRange Is Nothing Then
                        If numberBookmark.Range.Start = numberRange.Start And _
                           numberBookmark.Range.End = numberRange.End Then
                            bookmarkName = numberBookmark.Name
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next numberBookmark
    If Len(bookmarkName) = 0 Or numberRange Is Nothing Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The numbered image Equation invariant static number is missing."
    End If
    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    Set sequenceField = VTEquationSequenceFieldForBookmark( _
        documentObject, sequenceBookmarkName)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The numbered image Equation invariant SEQ is missing."
    End If
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If Not VTHelperParagraphOwnsNativeEquationSequence( _
       helperParagraph) Or helperParagraph.Start < paragraphRange.End Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The numbered image Equation invariant helper is invalid."
    End If

    Set formulaStartProbe = formulaRange.Duplicate
    formulaStartProbe.Collapse wdCollapseStart
    Set formulaEndProbe = formulaRange.Duplicate
    formulaEndProbe.Collapse wdCollapseEnd
    Set numberEndProbe = numberRange.Duplicate
    numberEndProbe.Collapse wdCollapseEnd
    documentObject.Repaginate

    VTNumberedEquationInvariantSnapshot = _
        "paragraph=" & CStr(paragraphRange.Start) & ":" & _
            CStr(paragraphRange.End) & _
        "|text=" & Replace$(Replace$(paragraphRange.Text, vbTab, "<TAB>"), vbCr, "<CR>") & _
        "|visibleStatic=" & CStr(numberRange.Start) & ":" & _
            CStr(numberRange.End) & ":" & numberRange.Text & _
        "|externalSeq=" & CStr(VTEquationFieldStart(sequenceField)) & _
            ":" & CStr(VTEquationFieldEnd(sequenceField)) & _
            ":" & sequenceField.Result.Text & ":" & _
                Trim$(sequenceField.Code.Text) & _
        "|helper=" & CStr(helperParagraph.Start) & ":" & _
            CStr(helperParagraph.End)
    VTNumberedEquationInvariantSnapshot = VTNumberedEquationInvariantSnapshot & _
        "|formula=" & CStr(formulaRange.Start) & ":" & _
            CStr(formulaRange.End) & ":" & CStr(formulaType) & _
            ":" & CStr(formulaRange.Font.Size) & _
            ":" & CStr(formulaRange.Font.Position) & _
        "|tabs=" & CStr(paragraphRange.ParagraphFormat.TabStops.Count) & _
            ":" & CStr(paragraphRange.ParagraphFormat.TabStops(1).Alignment) & _
            ":" & CStr(paragraphRange.ParagraphFormat.TabStops(1).Position) & _
            ":" & CStr(paragraphRange.ParagraphFormat.TabStops(2).Alignment) & _
            ":" & CStr(paragraphRange.ParagraphFormat.TabStops(2).Position)
    VTNumberedEquationInvariantSnapshot = VTNumberedEquationInvariantSnapshot & _
        "|xy=" & CStr(formulaStartProbe.Information( _
            wdHorizontalPositionRelativeToTextBoundary)) & _
            ":" & CStr(formulaEndProbe.Information( _
            wdHorizontalPositionRelativeToTextBoundary)) & _
            ":" & CStr(numberEndProbe.Information( _
            wdHorizontalPositionRelativeToTextBoundary)) & _
        "|line=" & CStr(formulaStartProbe.Information( _
            wdFirstCharacterLineNumber)) & ":" & _
            CStr(numberEndProbe.Information(wdFirstCharacterLineNumber))
    VTNumberedEquationInvariantSnapshot = VTNumberedEquationInvariantSnapshot & _
        "|bookmark=" & bookmarkName & ":" & numberRange.Text
End Function

Private Function VTNumberedDisplayTableInvariantSnapshot( _
    ByVal layoutTable As Table) As String

    Dim documentObject As Document
    Dim formulaRange As Range
    Dim numberRange As Range
    Dim mirrorRange As Range
    Dim sequenceRange As Range
    Dim formulaStartProbe As Range
    Dim formulaEndProbe As Range
    Dim numberProbe As Range
    Dim candidateBookmark As Bookmark
    Dim candidateField As Field
    Dim sequenceField As Field
    Dim referenceField As Field
    Dim numberBookmarkName As String
    Dim sequenceBookmarkName As String
    Dim suffixText As String
    Dim cellXml As String
    Dim formulaType As Long
    Dim centerParagraphCount As Long
    Dim mathParaPresent As Long
    Dim tailFontSize As Single
    Dim tailLineSpacing As Single

    If layoutTable Is Nothing Or layoutTable.Rows.Count <> 1 Or _
       layoutTable.Columns.Count <> 3 Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The numbered Equation invariant table is invalid."
    End If
    Set documentObject = layoutTable.Range.Document
    If layoutTable.Cell(1, 2).Range.OMaths.Count = 1 Then
        Set formulaRange = _
            VTNativeMathRoot(layoutTable.Cell(1, 2).Range.OMaths(1)).Range.Duplicate
        formulaType = layoutTable.Cell(1, 2).Range.OMaths(1).Type
    ElseIf layoutTable.Cell(1, 2).Range.InlineShapes.Count = 1 Then
        Set formulaRange = _
            layoutTable.Cell(1, 2).Range.InlineShapes(1).Range.Duplicate
        formulaType = -1
    Else
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The numbered Equation invariant formula is ambiguous."
    End If
    centerParagraphCount = _
        layoutTable.Cell(1, 2).Range.Paragraphs.Count
    If formulaType <> -1 Then
        cellXml = layoutTable.Cell(1, 2).Range.WordOpenXML
        If InStr(1, cellXml, "<m:oMathPara", vbBinaryCompare) > 0 Then
            mathParaPresent = 1
        End If
        If centerParagraphCount <> 2 Or mathParaPresent <> 1 Then
            Err.Raise vbObjectError + 7547, "VisualTeX", _
                "The numbered native invariant lost its display-cell structure."
        End If
        tailFontSize = _
            layoutTable.Cell(1, 2).Range.Paragraphs(2).Range.Font.Size
        tailLineSpacing = _
            layoutTable.Cell(1, 2).Range.Paragraphs(2).Range.ParagraphFormat.LineSpacing
    End If
    For Each candidateField In layoutTable.Cell(1, 3).Range.Fields
        If candidateField.Type = wdFieldRef Then
            sequenceBookmarkName = VTBookmarkNameInFieldCode( _
                documentObject, candidateField.Code.Text, _
                VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)
            If Len(sequenceBookmarkName) > 0 Then
                Set referenceField = candidateField
                Exit For
            End If
        End If
    Next candidateField
    If referenceField Is Nothing Or Len(sequenceBookmarkName) = 0 Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The numbered Equation invariant native REF is missing."
    End If
    suffixText = Mid$(sequenceBookmarkName, _
        Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX) + 1)
    numberBookmarkName = VT_WORD_NUMBER_BOOKMARK_PREFIX & suffixText
    If Not documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The numbered Equation invariant sequence Bookmark is missing."
    End If
    If Not documentObject.Bookmarks.Exists(numberBookmarkName) Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The numbered Equation invariant visible Bookmark is missing."
    End If
    Set mirrorRange = documentObject.Bookmarks( _
        numberBookmarkName).Range.Duplicate
    Set numberRange = mirrorRange.Duplicate
    Set sequenceRange = documentObject.Bookmarks( _
        sequenceBookmarkName).Range.Duplicate
    For Each candidateField In documentObject.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            If candidateField.Result.Start <= sequenceRange.Start And _
               candidateField.Result.End >= sequenceRange.End Then
                Set sequenceField = candidateField
                Exit For
            End If
        End If
    Next candidateField
    If sequenceField Is Nothing Or referenceField Is Nothing Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The numbered Equation invariant fields are missing."
    End If

    documentObject.Repaginate
    Set formulaStartProbe = formulaRange.Duplicate
    formulaStartProbe.Collapse wdCollapseStart
    Set formulaEndProbe = formulaRange.Duplicate
    formulaEndProbe.Collapse wdCollapseEnd
    Set numberProbe = numberRange.Duplicate
    numberProbe.Collapse wdCollapseEnd
    VTNumberedDisplayTableInvariantSnapshot = _
        "table=" & CStr(layoutTable.Range.Start) & ":" & _
            CStr(layoutTable.Range.End) & _
        "|columns=" & CStr(layoutTable.Columns(1).PreferredWidth) & _
            ":" & CStr(layoutTable.Columns(2).PreferredWidth) & _
            ":" & CStr(layoutTable.Columns(3).PreferredWidth) & _
        "|formula=" & CStr(formulaRange.Start) & ":" & _
            CStr(formulaRange.End) & ":" & CStr(formulaType) & _
            ":" & CStr(formulaRange.Font.Size) & _
            ":" & CStr(formulaRange.Font.Position) & _
        "|number=" & numberBookmarkName & ":" & _
            CStr(numberRange.Start) & ":" & CStr(numberRange.End) & _
            ":" & numberRange.Text & ":" & _
            Trim$(referenceField.Code.Text) & _
            ":mirror=" & mirrorRange.Text
    VTNumberedDisplayTableInvariantSnapshot = _
        VTNumberedDisplayTableInvariantSnapshot & _
        "|sequence=" & sequenceBookmarkName & ":" & _
        CStr(sequenceRange.Start) & ":" & CStr(sequenceRange.End)
    VTNumberedDisplayTableInvariantSnapshot = _
        VTNumberedDisplayTableInvariantSnapshot & ":" & _
        sequenceField.Result.Text & ":" & Trim$(sequenceField.Code.Text)
    VTNumberedDisplayTableInvariantSnapshot = _
        VTNumberedDisplayTableInvariantSnapshot & "|paragraphs=" & _
        CStr(layoutTable.Cell(1, 2).Range.Paragraphs(1).Range.Start) & _
        ":" & _
        CStr(layoutTable.Cell(1, 2).Range.Paragraphs(1).Range.End)
    VTNumberedDisplayTableInvariantSnapshot = _
        VTNumberedDisplayTableInvariantSnapshot & ":" & _
        CStr(layoutTable.Cell(1, 3).Range.Paragraphs(1).Range.Start) & _
        ":" & _
        CStr(layoutTable.Cell(1, 3).Range.Paragraphs(1).Range.End)
    VTNumberedDisplayTableInvariantSnapshot = _
        VTNumberedDisplayTableInvariantSnapshot & "|tabs=" & _
        CStr(VTCustomTabStopCount( _
            layoutTable.Cell(1, 2).Range.Paragraphs(1).Range)) & ":" & _
        CStr(VTCustomTabStopCount( _
            layoutTable.Cell(1, 3).Range.Paragraphs(1).Range))
    VTNumberedDisplayTableInvariantSnapshot = _
        VTNumberedDisplayTableInvariantSnapshot & _
        "|centerStructure=" & CStr(centerParagraphCount) & ":" & _
        CStr(mathParaPresent) & ":" & CStr(tailFontSize) & ":" & _
        CStr(tailLineSpacing)
    VTNumberedDisplayTableInvariantSnapshot = _
        VTNumberedDisplayTableInvariantSnapshot & _
        "|xy=" & CStr(formulaStartProbe.Information( _
            wdHorizontalPositionRelativeToTextBoundary)) & _
            ":" & CStr(formulaEndProbe.Information( _
            wdHorizontalPositionRelativeToTextBoundary)) & _
            ":" & CStr(numberProbe.Information( _
            wdHorizontalPositionRelativeToTextBoundary)) & _
        "|line=" & CStr(formulaStartProbe.Information( _
            wdFirstCharacterLineNumber)) & ":" & _
            CStr(numberProbe.Information(wdFirstCharacterLineNumber))
End Function

Private Sub VTAssertNumberedEquationInvariant( _
    ByVal documentObject As Document, _
    ByVal paragraphStart As Long, _
    ByVal expectedSnapshot As String, _
    ByVal assertionName As String)

    Dim actualSnapshot As String

    actualSnapshot = VTNumberedEquationInvariantSnapshot( _
        documentObject, paragraphStart)
    If actualSnapshot <> expectedSnapshot Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            assertionName & ": inserting a later display formula changed an earlier number" & _
            " [before=" & expectedSnapshot & "; after=" & actualSnapshot & "]."
    End If
End Sub

Private Function VTWrapNativeDisplayParagraphInTable( _
    ByVal formulaRange As Range) As Table

    Dim documentObject As Document
    Dim nativeEquation As OMath
    Dim paragraphRange As Range
    Dim sourceParagraph As Range
    Dim beforeRange As Range
    Dim afterRange As Range
    Dim insertionRange As Range
    Dim centerRange As Range
    Dim trailingParagraph As Range
    Dim layoutTable As Table
    Dim paragraphStart As Long
    Dim cellXml As String
    Dim operationStage As String
    Dim operationErrorNumber As Long
    Dim operationErrorDescription As String

    On Error GoTo WrapFailed
    operationStage = "validate-input"
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The native display paragraph Range is missing."
    End If
    If formulaRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The native display paragraph has no unique OMath."
    End If
    If formulaRange.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The native display paragraph is already inside a table."
    End If

    Set documentObject = formulaRange.Document
    Set nativeEquation = formulaRange.OMaths(1)
    Set paragraphRange = nativeEquation.Range.Paragraphs(1).Range.Duplicate
    paragraphStart = paragraphRange.Start
    Set beforeRange = paragraphRange.Duplicate
    beforeRange.End = nativeEquation.Range.Start
    Set afterRange = paragraphRange.Duplicate
    afterRange.Start = nativeEquation.Range.End
    If VTWordRangeHasMeaningfulText(beforeRange) Or _
       VTWordRangeHasMeaningfulText(afterRange) Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "A numbered display OMath must occupy its own paragraph."
    End If

    ' Word for Mac rejects table creation when the insertion anchor is already
    ' attached to a display-math paragraph. Force the source OMath inline and
    ' re-resolve its paragraph before creating the empty table. Afterward,
    ' promote the source paragraph to true display math and transfer that
    ' complete paragraph within the same document.
    operationStage = "prepare-source-inline"
    nativeEquation.Type = wdOMathInline
    nativeEquation.Range.Font.Position = 0
    nativeEquation.Range.Font.Size = _
        VTPreferredNativeDisplayFontSize(nativeEquation.Range)
    Set paragraphRange = documentObject.Range( _
        Start:=paragraphStart, End:=paragraphStart).Paragraphs(1).Range.Duplicate
    If paragraphRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word lost the source OMath while preparing table creation."
    End If
    Set nativeEquation = paragraphRange.OMaths(1)
    If nativeEquation.Type <> wdOMathInline Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word did not return the source OMath to inline mode."
    End If
    Set beforeRange = paragraphRange.Duplicate
    beforeRange.End = nativeEquation.Range.Start
    Set afterRange = paragraphRange.Duplicate
    afterRange.Start = nativeEquation.Range.End
    If VTWordRangeHasMeaningfulText(beforeRange) Or _
       VTWordRangeHasMeaningfulText(afterRange) Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word added unexpected text while preparing table creation."
    End If

    operationStage = "create-empty-table"
    Set insertionRange = paragraphRange.Duplicate
    insertionRange.Collapse wdCollapseEnd
    Set layoutTable = documentObject.Tables.Add( _
        Range:=insertionRange, NumRows:=1, NumColumns:=3)
    If layoutTable.Rows.Count <> 1 Or layoutTable.Columns.Count <> 3 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word could not create the numbered display table."
    End If

    operationStage = "promote-source-display"
    Set paragraphRange = documentObject.Range( _
        Start:=paragraphStart, End:=paragraphStart).Paragraphs(1).Range.Duplicate
    If paragraphRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word lost the source OMath after creating the empty table."
    End If
    Set nativeEquation = paragraphRange.OMaths(1)
    nativeEquation.Type = wdOMathDisplay
    nativeEquation.Justification = wdOMathJcCenter
    nativeEquation.Range.Font.Position = 0
    nativeEquation.Range.Font.Size = _
        VTPreferredNativeDisplayFontSize(nativeEquation.Range)
    nativeEquation.BuildUp

    operationStage = "verify-source-display"
    Set paragraphRange = documentObject.Range( _
        Start:=paragraphStart, End:=paragraphStart).Paragraphs(1).Range.Duplicate
    If paragraphRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word lost the unique native OMath while building display math."
    End If
    Set nativeEquation = paragraphRange.OMaths(1)
    If nativeEquation.Type <> wdOMathDisplay Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word did not preserve the source equation as display math."
    End If
    Set beforeRange = paragraphRange.Duplicate
    beforeRange.End = nativeEquation.Range.Start
    Set afterRange = paragraphRange.Duplicate
    afterRange.Start = nativeEquation.Range.End
    If VTWordRangeHasMeaningfulText(beforeRange) Or _
       VTWordRangeHasMeaningfulText(afterRange) Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word added unexpected text while building display math."
    End If

    operationStage = "transfer-display-paragraph"
    Set sourceParagraph = documentObject.Range( _
        Start:=paragraphStart, End:=paragraphStart).Paragraphs(1).Range.Duplicate
    If sourceParagraph.OMaths.Count <> 1 Or _
       sourceParagraph.OMaths(1).Type <> wdOMathDisplay Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word could not re-resolve the source display paragraph after table creation."
    End If
    Set centerRange = layoutTable.Cell(1, 2).Range.Duplicate
    centerRange.End = centerRange.End - 1
    centerRange.FormattedText = sourceParagraph.FormattedText

    operationStage = "verify-transferred-display"
    If layoutTable.Cell(1, 2).Range.Paragraphs.Count <> 2 Or _
       layoutTable.Cell(1, 2).Range.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word did not retain the display paragraph and required cell tail."
    End If
    Set nativeEquation = layoutTable.Cell(1, 2).Range.OMaths(1)
    If nativeEquation.Type <> wdOMathDisplay Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Same-document transfer downgraded the center-cell equation."
    End If
    Set trailingParagraph = _
        layoutTable.Cell(1, 2).Range.Paragraphs(2).Range.Duplicate
    If trailingParagraph.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(trailingParagraph) Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The required display-cell tail contains unexpected content."
    End If

    operationStage = "delete-source-display"
    Set sourceParagraph = documentObject.Range( _
        Start:=paragraphStart, End:=paragraphStart).Paragraphs(1).Range.Duplicate
    If sourceParagraph.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word could not re-resolve the original display paragraph."
    End If
    sourceParagraph.Delete

    operationStage = "configure-table"
    VTConfigureNumberedDisplayTable layoutTable
    If layoutTable.Cell(1, 2).Range.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word lost the center-cell OMath after removing the source paragraph."
    End If
    operationStage = "finalize-center-display"
    Set nativeEquation = layoutTable.Cell(1, 2).Range.OMaths(1)
    nativeEquation.Type = wdOMathDisplay
    nativeEquation.Justification = wdOMathJcCenter
    nativeEquation.BuildUp
    If layoutTable.Cell(1, 2).Range.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word lost the center-cell OMath while rebuilding display math."
    End If
    Set nativeEquation = layoutTable.Cell(1, 2).Range.OMaths(1)
    nativeEquation.Range.Font.Position = 0
    nativeEquation.Range.Font.Size = _
        VTPreferredNativeDisplayFontSize(layoutTable.Cell(1, 2).Range)

    operationStage = "compact-cell-tail"
    VTCompactNativeDisplayCellTail layoutTable

    operationStage = "verify-cell-openxml"
    cellXml = layoutTable.Cell(1, 2).Range.WordOpenXML
    If InStr(1, cellXml, "<m:oMathPara", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The numbered display cell has no native m:oMathPara structure."
    End If
    If layoutTable.Cell(1, 2).Range.Paragraphs.Count <> 2 Or _
       layoutTable.Cell(1, 2).Range.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The compacted display cell lost its stable two-paragraph structure."
    End If
    Set VTWrapNativeDisplayParagraphInTable = layoutTable
    Exit Function

WrapFailed:
    operationErrorNumber = Err.Number
    operationErrorDescription = Err.Description
    Err.Raise operationErrorNumber, _
        "VisualTeX native display table", _
        "VTWrapNativeDisplayParagraphInTable/" & operationStage & ": " & _
        operationErrorDescription
End Function

Private Sub VTCompactNativeDisplayCellTail(ByVal layoutTable As Table)
    Dim trailingParagraph As Range

    If layoutTable Is Nothing Or layoutTable.Rows.Count <> 1 Or _
       layoutTable.Columns.Count <> 3 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The native display tail requires a one-row three-column table."
    End If
    If layoutTable.Cell(1, 2).Range.Paragraphs.Count <> 2 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The native display cell has no unique required tail paragraph."
    End If

    Set trailingParagraph = _
        layoutTable.Cell(1, 2).Range.Paragraphs(2).Range.Duplicate
    If trailingParagraph.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(trailingParagraph) Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The native display tail contains unexpected content."
    End If

    With trailingParagraph
        .Font.Position = 0
        .Font.Size = 1!
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LeftIndent = 0!
            .RightIndent = 0!
            .FirstLineIndent = 0!
            .SpaceBefore = 0!
            .SpaceAfter = 0!
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 1!
        End With
    End With
End Sub

Private Sub VTAssertBodyEquationReferenceVisible( _
    ByVal documentObject As Document, _
    ByVal referenceField As Field, _
    ByVal assertionName As String)

    Dim visibleSize As Single

    If documentObject Is Nothing Or referenceField Is Nothing Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            assertionName & ": the body Equation reference is missing."
    End If
    visibleSize = VTVisibleEquationNumberFontSize(documentObject)
    If referenceField.Type <> wdFieldRef Or _
       referenceField.Result.Information(wdWithInTable) Or _
       referenceField.Result.Font.Hidden <> False Or _
       referenceField.Result.Font.Color <> wdColorAutomatic Or _
       Abs(referenceField.Result.Font.Size - visibleSize) > 0.1 Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            assertionName & ": the body Equation reference is not visibly" & _
            " formatted [result=" & referenceField.Result.Text & _
            "; size=" & CStr(referenceField.Result.Font.Size) & _
            "; hidden=" & CStr(referenceField.Result.Font.Hidden) & _
            "; color=" & CStr(referenceField.Result.Font.Color) & _
            "; expectedSize=" & CStr(visibleSize) & "]."
    End If
End Sub

Private Function VTCustomTabStopCount( _
    ByVal paragraphRange As Range) As Long

    Dim tabStop As TabStop
    Dim customCount As Long

    If paragraphRange Is Nothing Then Exit Function
    For Each tabStop In paragraphRange.ParagraphFormat.TabStops
        If tabStop.CustomTab Then customCount = customCount + 1
    Next tabStop
    VTCustomTabStopCount = customCount
End Function

Public Sub VisualTeX_RunWordChineseImageAlignmentRegression()
    Dim testDocument As Document
    Dim shape As InlineShape
    Dim paragraphRange As Range
    Dim originalMode As Long
    Dim modes As Variant
    Dim mode As Variant
    Dim imageIndex As Long
    Dim originalHeight As Double
    Dim originalWidth As Double
    Dim report As String
    Dim resultPath As String
    Dim failureNumber As Long
    Dim failureDescription As String
    resultPath = VTApplicationSupportRoot() & "/Tests/word-chinese-image-alignment-result.txt"
    On Error GoTo RegressionFailed
    Set testDocument = ActiveDocument
    If testDocument.Name <> "word-chinese-image-alignment-fixture.docx" Or _
       testDocument.InlineShapes.Count <> 45 Then _
        Err.Raise 5, "VisualTeX regression", "Open the 45-image reference fixture."
    modes = Array(wdBaselineAlignBaseline, wdBaselineAlignTop, _
        wdBaselineAlignCenter, wdBaselineAlignFarEast50, wdBaselineAlignAuto)
    VT_WORD_TYPOGRAPHY_REGRESSION_RUNNING = True
    report = "PASS" & vbLf
    For imageIndex = 1 To testDocument.InlineShapes.Count
        Set shape = testDocument.InlineShapes(imageIndex)
        Set paragraphRange = shape.Range.Paragraphs(1).Range.Duplicate
        originalMode = paragraphRange.ParagraphFormat.BaseLineAlignment
        originalHeight = VTWordImageGlyphHeight(shape)
        originalWidth = shape.Width
        ' The caret stays in ordinary paragraph text, not on the picture.
        ' Exercise the real selection/OnTime repair for every typography mode.
        testDocument.Range(paragraphRange.Start, paragraphRange.Start).Select
        For Each mode In modes
            paragraphRange.ParagraphFormat.BaseLineAlignment = CLng(mode)
            VisualTeX_WatchSelectedImageFormulaSize
            If paragraphRange.ParagraphFormat.BaseLineAlignment <> CLng(mode) Then _
                Err.Raise 5, "VisualTeX regression", "Paragraph typography was overwritten."
        Next mode
        paragraphRange.ParagraphFormat.BaseLineAlignment = originalMode
        VisualTeX_WatchSelectedImageFormulaSize
        If imageIndex <= 35 Then
            VTApplyWordImageFormulaFontSize shape, 24#
            VTApplyWordImageFormulaFontSize shape, 18#
        Else
            VTApplyWordImagePosition shape, VTExpectedImageFormulaPosition( _
                shape, VTVisibleEquationNumberFontSize(testDocument), "Chinese typography")
        End If
        If Abs(VTWordImageGlyphHeight(shape) - originalHeight) > 0.1 Or _
           Abs(shape.Width - originalWidth) > 0.1 Then _
            Err.Raise 5, "VisualTeX regression", "Resize round trip changed image geometry" & _
                " [before=" & CStr(originalWidth) & "x" & CStr(originalHeight) & _
                "; after=" & CStr(shape.Width) & "x" & CStr(shape.Height) & "]."
        If paragraphRange.ParagraphFormat.BaseLineAlignment <> originalMode Then Err.Raise 5
        report = report & "image=" & CStr(imageIndex) & "; alignment=" & _
            CStr(originalMode) & "; position=" & CStr(shape.Range.Font.Position) & vbLf
    Next imageIndex
    ' A mode/style change outside the caret paragraph must also be observed.
    Set shape = testDocument.InlineShapes(1)
    Set paragraphRange = shape.Range.Paragraphs(1).Range.Duplicate
    testDocument.Range(0, 0).Select
    paragraphRange.ParagraphFormat.BaseLineAlignment = wdBaselineAlignTop
    VT_WORD_TYPOGRAPHY_REGRESSION_RUNNING = False
    VisualTeX_WatchSelectedImageFormulaSize
    If shape.Range.Font.Position <> 18 Then _
        Err.Raise 5, "VisualTeX regression", "Document mode observer did not repair the remote paragraph."
    paragraphRange.ParagraphFormat.BaseLineAlignment = wdBaselineAlignBaseline
    VisualTeX_WatchSelectedImageFormulaSize
    VT_WORD_TYPOGRAPHY_REGRESSION_RUNNING = True
    ' A whole-document selection must repair all images across paragraphs.
    testDocument.Content.Select
    VisualTeX_WatchSelectedImageFormulaSize
    testDocument.Repaginate
    testDocument.Save
    VTWriteTextAtomic resultPath, report
    testDocument.Range(0, 0).Select
    VT_WORD_TYPOGRAPHY_REGRESSION_RUNNING = False
    VTEnsureImageSizeWatchScheduled
    Exit Sub
RegressionFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    VT_WORD_TYPOGRAPHY_REGRESSION_RUNNING = False
    VTWriteTextAtomic resultPath, "FAIL" & vbLf & "image=" & CStr(imageIndex) & _
        vbLf & CStr(failureNumber) & ":" & failureDescription
End Sub

Public Sub VisualTeX_RunWordUnnumberedImageParagraphMarkRegression()
    Const formulaId As String = _
        "73737373-7373-4737-8737-737373737373"

    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim insertionRange As Range
    Dim paragraphRange As Range
    Dim paragraphMarkRange As Range
    Dim fixtureRoot As String
    Dim resultPath As String
    Dim imagePath As String
    Dim imageFileName As String
    Dim expectedPosition As Long
    Dim actualPosition As Long
    Dim paragraphMarkPosition As Long
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    resultPath = fixtureRoot & _
        "/word-unnumbered-image-paragraph-mark-regression-result.txt"
    On Error GoTo RegressionFailed
    If Documents.Count > 0 Then
        Err.Raise vbObjectError + 7613, "VisualTeX regression", _
            "The unnumbered-image paragraph-mark regression requires Word" & _
            " to have no open documents."
    End If

    regressionStage = "resolve-real-image"
    imageFileName = Dir$(VTApplicationSupportRoot() & "/ImageDocuments/*.png")
    If Len(imageFileName) > 0 Then
        imagePath = VTApplicationSupportRoot() & "/ImageDocuments/" & imageFileName
    Else
        imagePath = VTPlaceholderImagePath()
    End If

    regressionStage = "create-visible-fixture"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.Activate
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.ActiveWindow.View.ShowAll = True
    testDocument.ActiveWindow.View.Zoom.Percentage = 150
    testDocument.Styles(wdStyleNormal).Font.Size = 11!
    testDocument.Content.Text = _
        "Unnumbered image display: centered paragraph mark" & vbCr
    Set insertionRange = testDocument.Range( _
        Start:=testDocument.Content.End - 1, _
        End:=testDocument.Content.End - 1)
    Set formulaShape = testDocument.InlineShapes.AddPicture( _
        FileName:=imagePath, LinkToFile:=False, _
        SaveWithDocument:=True, Range:=insertionRange)
    formulaShape.LockAspectRatio = msoTrue
    formulaShape.Height = 40!
    formulaShape.Title = VTFormulaReference(formulaId, "block", False)
    VTSetWordFormulaFormat testDocument, formulaId, "block", False
    VTSetWordImageInkCenterYRatio testDocument, formulaId, 0.5

    regressionStage = "normalize-unnumbered-display"
    VTNormalizeUnnumberedDisplayParagraph formulaShape.Range
    VTNormalizeImageDisplayParagraph formulaShape.Range
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    Set paragraphMarkRange = testDocument.Range( _
        Start:=paragraphRange.End - 1, End:=paragraphRange.End)
    expectedPosition = VTExpectedImageFormulaPosition( _
        formulaShape, CDbl(VTVisibleEquationNumberFontSize(testDocument)), _
        "The unnumbered image paragraph-mark regression")
    actualPosition = CLng(formulaShape.Range.Font.Position)
    paragraphMarkPosition = CLng(paragraphMarkRange.Font.Position)

    regressionStage = "verify-layout"
    If paragraphRange.ParagraphFormat.Alignment <> wdAlignParagraphCenter Or _
       VTCustomTabStopCount(paragraphRange) <> 0 Or _
       expectedPosition >= 0 Or actualPosition <> expectedPosition Or _
       paragraphMarkPosition <> 0 Then
        Err.Raise vbObjectError + 7613, "VisualTeX regression", _
            "The unnumbered image and paragraph mark do not share the" & _
            " expected visual-center baseline" & _
            " [alignment=" & CStr(paragraphRange.ParagraphFormat.Alignment) & _
            "; customTabs=" & CStr(VTCustomTabStopCount(paragraphRange)) & _
            "; formulaPosition=" & CStr(actualPosition) & _
            "; expectedPosition=" & CStr(expectedPosition) & _
            "; paragraphMarkPosition=" & CStr(paragraphMarkPosition) & "]."
    End If

    regressionStage = "write-result"
    VTWriteTextAtomic resultPath, _
        "PASS" & vbLf & _
        "formulaPosition=" & CStr(actualPosition) & vbLf & _
        "paragraphMarkPosition=" & CStr(paragraphMarkPosition) & vbLf & _
        "height=" & CStr(formulaShape.Height) & vbLf & _
        "visualInspection=visible-word-show-all" & vbLf
    testDocument.Range(Start:=0, End:=0).Select
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic resultPath, _
        "FAIL" & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), vbLf, " ") & vbLf
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX unnumbered image paragraph-mark regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Private Sub VTNativeParityTrace(ByVal stage As String, ByVal doc As Document, Optional ByVal exact As Range)
    If Not VT_NATIVE_PARITY_DIAGNOSTIC Then Exit Sub
    On Error Resume Next
    Dim value As String
    value = stage & "|doc=" & doc.Name & "|maths=" & CStr(doc.OMaths.Count)
    If Not exact Is Nothing Then value = value & "|rangeDoc=" & exact.Document.Name & "|range=" & CStr(exact.Start) & "-" & CStr(exact.End) & "|text=" & Replace$(exact.Text, vbCr, "[CR]")
    VTAppendText VTApplicationSupportRoot() & "/Tests/native-parity-debug.txt", value & vbLf
End Sub

Private Function VTTryResolveCopiedNativeFormulaSource( _
    ByVal selectedRange As Range, _
    ByRef copiedMath As OMath, _
    ByRef sourceDocument As Document, _
    ByRef sourceFormulaId As String, _
    ByRef encodedMetadata As String, _
    ByRef displayMode As String, _
    ByRef numbered As Boolean) As Boolean

    Dim destinationDocument As Document
    Dim candidateDocument As Document
    Dim candidateBookmark As Bookmark
    Dim candidateMath As OMath
    Dim existingBookmark As Bookmark
    Dim candidateFormulaId As String
    Dim candidateMetadata As String
    Dim candidateDisplayMode As String
    Dim candidateNumbered As Boolean
    Dim copiedDisplayMode As String
    Dim copiedSignature As String
    Dim localBookmarkFormulaId As String
    Dim localMetadata As String
    Dim localDisplayMode As String
    Dim localNumbered As Boolean
    Dim localBackingState As Boolean
    Dim candidateSignature As String
    Dim destinationDocumentId As String
    Dim candidateDocumentId As String
    Dim passIndex As Long
    Dim candidateDistance As Double
    Dim bestDistance As Double
    Dim foundMatch As Boolean
    Dim ambiguousCrossDocumentMatch As Boolean
    Dim bestMetadata As String
    Dim bestDisplayMode As String
    Dim bestNumbered As Boolean

    Set copiedMath = Nothing
    Set sourceDocument = Nothing
    sourceFormulaId = ""
    encodedMetadata = ""
    displayMode = ""
    numbered = False
    If selectedRange Is Nothing Then Exit Function
    If Not VTTryResolveSingleNativeMath(selectedRange, copiedMath) Then Exit Function

    Set destinationDocument = copiedMath.Range.Document
    destinationDocumentId = _
        VTWordDocumentIdentityForDocument(destinationDocument)
    If copiedMath.Type = wdOMathInline Then
        copiedDisplayMode = "inline"
    Else
        copiedDisplayMode = "block"
    End If
    copiedSignature = VTNativeMathCopySignature(copiedMath)
    If Len(copiedSignature) = 0 Then Exit Function

    ' Cross-document paste can preserve VT_F_<formulaId> because the target
    ' document has no Bookmark-name collision. A genuine managed formula has
    ' matching document Variables behind that Bookmark; a pasted external copy
    ' has the Bookmark but not those per-document Variables. In the latter case
    ' the Bookmark gives us an exact source formulaId, so prefer it over any
    ' content-signature heuristic.
    If VTTryFindNativeFormulaBookmarkLocally( _
       selectedRange, existingBookmark) Then
        If Not VTTryFormulaIdFromNativeBookmark( _
           existingBookmark.Name, localBookmarkFormulaId) Then Exit Function
        localMetadata = ""
        localDisplayMode = ""
        localNumbered = False
        On Error Resume Next
        localBackingState = _
            VTTryReadWordMetadataPayload( _
                destinationDocument, localBookmarkFormulaId, _
                localMetadata) And _
            VTTryReadWordFormulaFormat( _
                destinationDocument, localBookmarkFormulaId, _
                localDisplayMode, localNumbered)
        Err.Clear
        On Error GoTo 0
        If localBackingState Then Exit Function

        foundMatch = False
        ambiguousCrossDocumentMatch = False
        For Each candidateDocument In Documents
            candidateDocumentId = _
                VTWordDocumentIdentityForDocument(candidateDocument)
            If candidateDocumentId <> destinationDocumentId Then
                candidateFormulaId = localBookmarkFormulaId
                If candidateDocument.Bookmarks.Exists( _
                   VTNativeFormulaBookmarkName(candidateFormulaId)) Then
                    Set candidateBookmark = candidateDocument.Bookmarks( _
                        VTNativeFormulaBookmarkName(candidateFormulaId))
                    Set candidateMath = VTNativeMathForBookmark(candidateBookmark)
                    If Not candidateMath Is Nothing Then
                        candidateMetadata = ""
                        candidateDisplayMode = ""
                        candidateNumbered = False
                        If VTTryReadWordMetadataPayload( _
                           candidateDocument, candidateFormulaId, _
                           candidateMetadata) And _
                           VTTryReadWordFormulaFormat( _
                               candidateDocument, candidateFormulaId, _
                               candidateDisplayMode, candidateNumbered) Then
                            candidateSignature = _
                                VTNativeMathCopySignature(candidateMath)
                            If Len(candidateSignature) > 0 And _
                               StrComp(candidateSignature, copiedSignature, _
                                   vbBinaryCompare) = 0 Then
                                If Not foundMatch Then
                                    foundMatch = True
                                    Set sourceDocument = candidateDocument
                                    sourceFormulaId = candidateFormulaId
                                    bestMetadata = candidateMetadata
                                    If candidateNumbered Then
                                        bestDisplayMode = "block"
                                    Else
                                        bestDisplayMode = copiedDisplayMode
                                    End If
                                    bestNumbered = candidateNumbered
                                ElseIf StrComp( _
                                   bestMetadata, candidateMetadata, _
                                   vbBinaryCompare) <> 0 Or _
                                   bestNumbered <> candidateNumbered Then
                                    ambiguousCrossDocumentMatch = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next candidateDocument
        If foundMatch And Not ambiguousCrossDocumentMatch Then
            encodedMetadata = bestMetadata
            displayMode = bestDisplayMode
            numbered = bestNumbered
            VTTryResolveCopiedNativeFormulaSource = True
        End If
        Exit Function
    End If

    ' The common case is a paste inside the same Word document. Resolve it
    ' directly from that document's Bookmark collection instead of relying on
    ' the global Documents enumerator: Word for Mac can expose hidden/temporary
    ' documents through different COM proxy collections even though direct
    ' access to copiedMath.Range.Document is stable.
    bestDistance = 2147483647#
    foundMatch = False
    For Each candidateBookmark In destinationDocument.Bookmarks
        candidateFormulaId = ""
        If VTTryFormulaIdFromNativeBookmark( _
           candidateBookmark.Name, candidateFormulaId) Then
            Set candidateMath = VTNativeMathForBookmark(candidateBookmark)
            If Not candidateMath Is Nothing Then
                candidateMetadata = ""
                candidateDisplayMode = ""
                candidateNumbered = False
                If VTTryReadWordMetadataPayload( _
                   destinationDocument, candidateFormulaId, _
                   candidateMetadata) And _
                   VTTryReadWordFormulaFormat( _
                       destinationDocument, candidateFormulaId, _
                       candidateDisplayMode, candidateNumbered) Then
                    candidateSignature = _
                        VTNativeMathCopySignature(candidateMath)
                    If Len(candidateSignature) > 0 And _
                       StrComp(candidateSignature, copiedSignature, _
                           vbBinaryCompare) = 0 Then
                        candidateDistance = Abs( _
                            CDbl(candidateMath.Range.Start) - _
                            CDbl(copiedMath.Range.Start))
                        If Not foundMatch Or candidateDistance < bestDistance Then
                            foundMatch = True
                            bestDistance = candidateDistance
                            Set sourceDocument = destinationDocument
                            sourceFormulaId = candidateFormulaId
                            bestMetadata = candidateMetadata
                            If candidateNumbered Then
                                bestDisplayMode = "block"
                            Else
                                bestDisplayMode = copiedDisplayMode
                            End If
                            bestNumbered = candidateNumbered
                        End If
                    End If
                End If
            End If
        End If
    Next candidateBookmark
    If foundMatch Then
        encodedMetadata = bestMetadata
        displayMode = bestDisplayMode
        numbered = bestNumbered
        VTTryResolveCopiedNativeFormulaSource = True
        Exit Function
    End If

    ' Word deliberately drops the unique VT_F_ Bookmark when a native OMath is
    ' copied. Recover the source from the persisted structural signature. Prefer
    ' a source in the destination document, choosing the nearest exact match;
    ' then fall back to other currently open Word documents for cross-document
    ' paste. Cross-document matches are accepted only when their metadata and
    ' display format agree, so unrelated look-alike formulas are never merged.
    For passIndex = 0 To 1
        foundMatch = False
        ambiguousCrossDocumentMatch = False
        bestDistance = 2147483647#
        For Each candidateDocument In Documents
            candidateDocumentId = _
                VTWordDocumentIdentityForDocument(candidateDocument)
            If (passIndex = 0 And _
                candidateDocumentId = destinationDocumentId) Or _
               (passIndex = 1 And _
                candidateDocumentId <> destinationDocumentId) Then
                For Each candidateBookmark In candidateDocument.Bookmarks
                    candidateFormulaId = ""
                    If VTTryFormulaIdFromNativeBookmark( _
                       candidateBookmark.Name, candidateFormulaId) Then
                        Set candidateMath = VTNativeMathForBookmark(candidateBookmark)
                        If Not candidateMath Is Nothing Then
                            candidateMetadata = ""
                            candidateDisplayMode = ""
                            candidateNumbered = False
                            If VTTryReadWordMetadataPayload( _
                               candidateDocument, candidateFormulaId, _
                               candidateMetadata) And _
                               VTTryReadWordFormulaFormat( _
                                   candidateDocument, candidateFormulaId, _
                                   candidateDisplayMode, candidateNumbered) Then
                                candidateSignature = _
                                    VTNativeMathCopySignature(candidateMath)
                                If Len(candidateSignature) > 0 And _
                                   StrComp(candidateSignature, copiedSignature, _
                                       vbBinaryCompare) = 0 Then
                                    If passIndex = 0 Then
                                        candidateDistance = Abs( _
                                            CDbl(candidateMath.Range.Start) - _
                                            CDbl(copiedMath.Range.Start))
                                        If Not foundMatch Or _
                                           candidateDistance < bestDistance Then
                                            foundMatch = True
                                            bestDistance = candidateDistance
                                            Set sourceDocument = candidateDocument
                                            sourceFormulaId = candidateFormulaId
                                            bestMetadata = candidateMetadata
                                            If candidateNumbered Then
                                                bestDisplayMode = "block"
                                            Else
                                                bestDisplayMode = copiedDisplayMode
                                            End If
                                            bestNumbered = candidateNumbered
                                        End If
                                    ElseIf Not foundMatch Then
                                        foundMatch = True
                                        Set sourceDocument = candidateDocument
                                        sourceFormulaId = candidateFormulaId
                                        bestMetadata = candidateMetadata
                                        If candidateNumbered Then
                                            bestDisplayMode = "block"
                                        Else
                                            bestDisplayMode = copiedDisplayMode
                                        End If
                                        bestNumbered = candidateNumbered
                                    ElseIf StrComp( _
                                       bestMetadata, candidateMetadata, _
                                       vbBinaryCompare) <> 0 Or _
                                       bestDisplayMode <> candidateDisplayMode Or _
                                       bestNumbered <> candidateNumbered Then
                                        ambiguousCrossDocumentMatch = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next candidateBookmark
            End If
        Next candidateDocument
        If foundMatch Then
            If ambiguousCrossDocumentMatch Then
                Set sourceDocument = Nothing
                sourceFormulaId = ""
                Exit Function
            End If
            encodedMetadata = bestMetadata
            displayMode = bestDisplayMode
            numbered = bestNumbered
            VTTryResolveCopiedNativeFormulaSource = True
            Exit Function
        End If
    Next passIndex
End Function

Private Function VTNativeMathCopySignature( _
    ByVal nativeMath As OMath) As String

    Dim signatureRange As Range
    Dim formulaContentRange As Range
    Dim functionIndex As Long
    Dim functionTypes As String
    Dim signatureText As String

    If nativeMath Is Nothing Then Exit Function
    Set signatureRange = nativeMath.Range.Duplicate

    ' A copied numbered Equation can carry the visible array marker while Word
    ' drops the unique VT_F_ Bookmark. Strip the generated marker whenever the
    ' array shape is recognizable, regardless of whether the copied object still
    ' owns any VisualTeX numbering Bookmarks.
    On Error Resume Next
    Set formulaContentRange = _
        VTNativeEquationFormulaContentRange(signatureRange)
    Err.Clear
    On Error GoTo 0
    If Not formulaContentRange Is Nothing Then
        Set signatureRange = formulaContentRange.Duplicate
    End If

    signatureText = signatureRange.Text
    signatureText = Replace$(signatureText, vbCrLf, vbLf)
    signatureText = Replace$(signatureText, vbCr, vbLf)
    signatureText = Replace$(signatureText, Chr$(7), "")

    On Error Resume Next
    For functionIndex = 1 To nativeMath.Functions.Count
        functionTypes = functionTypes & "," & _
            CStr(nativeMath.Functions(functionIndex).Type)
    Next functionIndex
    Err.Clear
    On Error GoTo 0

    ' Copy/Paste can legitimately change wdOMathInline <-> wdOMathDisplay when
    ' Word moves the same equation between paragraph contexts. The copy identity
    ' signature therefore excludes OMath.Type and formula-specific numbering
    ' state while retaining the mathematical text and ordered function kinds.
    VTNativeMathCopySignature = "copy-v1:" & VTStableTextHash( _
        CStr(signatureRange.OMaths.Count) & "|" & _
        functionTypes & "|" & signatureText)
End Function

Public Sub VisualTeX_RunWordNumberingEditRegression()
    Dim doc As Document
    Dim dispatch As Object
    Dim shape As InlineShape
    Dim equation As Range
    Dim insertion As Range
    Dim result As Range
    Dim targets As Collection
    Dim ids As New Collection
    Dim native As Boolean
    Dim chapter As Boolean
    Dim representation As Long
    Dim formatIndex As Long
    Dim index As Long
    Dim position As Variant
    Dim sample As Long
    Dim numbered As Boolean
    Dim id As String
    Dim root As String
    Dim outputPath As String
    Dim stage As String
    Dim report As String
    Dim started As Single
    Dim elapsed As Double
    Dim created As Boolean
    Dim expectedCount As Long
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed
    root = VTApplicationSupportRoot() & "/Tests/numbering-edit"
    For representation = 0 To 1
        native = representation = 1
        For formatIndex = 0 To 1
            chapter = formatIndex = 1
            stage = "create-fixture"
            VTWriteTextAtomic root & "/progress.txt", CStr(native) & "|" & CStr(chapter) & "|" & stage
            outputPath = Left$(root, InStr(1, root, "/Library/", vbBinaryCompare) - 1) & _
                "/Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch/numbering-acceptance-" & _
                CStr(representation) & "-" & CStr(formatIndex) & ".docx"
            Set doc = Documents.Open(FileName:=outputPath, ReadOnly:=False, AddToRecentFiles:=False)
            Application.ScreenUpdating = False
            doc.Content.Text = "BODY_BEFORE" & vbCr
            doc.Paragraphs(1).Range.Style = wdStyleHeading1
            If chapter Then
                VTSetEquationNumberingFormat doc, VT_WORD_NUMBERING_MODE_CHAPTER, "."
            Else
                VTSetEquationNumberingFormat doc, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
            End If
            For index = 1 To 50
                id = "91919191-9191-4191-8191-" & Format$(index, "000000000000")
                Set dispatch = VTRegressionNumberingEditDispatch(doc, id, native, True)
                Set insertion = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                If native Then
                    Set equation = VTInsertNativeEquationAtRange(insertion, _
                        CStr(dispatch("ommlBase64")), CStr(dispatch("nativeDocumentPath")), "inline", True, False)
                    VTSetNativeFormulaBookmark doc, equation, id
                    Set result = VTEnsureNativeEquationNumber(equation, 30#, id, "", created)
                    Set equation = doc.Bookmarks(VTNativeFormulaBookmarkName(id)).Range.Duplicate
                    equation.Font.Size = 14!
                Else
                    Set shape = VTAddWordFormulaPicture(doc, insertion, _
                        CStr(dispatch("vectorDocumentPath")), CStr(dispatch("fallbackImagePath")))
                    shape.Width = CSng(dispatch("widthPoints"))
                    shape.Height = CSng(dispatch("heightPoints"))
                    shape.Title = VTFormulaReference(id, "block", True)
                    shape.AlternativeText = CStr(dispatch("metadata"))
                    Set result = VTEnsureImageEquationNumber(shape, shape.Height, id, "", created)
                    VTSetWordImageOwnerBookmark doc, shape, id
                    Set equation = shape.Range.Duplicate
                End If
                VTSetWordLatexPayload doc, id, CStr(dispatch("latexBase64"))
                VTSetWordOmmlPayload doc, id, CStr(dispatch("ommlBase64"))
                VTSetWordMetadataPayload doc, id, CStr(dispatch("metadata"))
                VTSetWordFormulaFormat doc, id, "block", True
                VTSetWordImageScaleState doc, id, 14#, CDbl(dispatch("referenceWidthPt")), _
                    CDbl(dispatch("referenceHeightPt")), CDbl(dispatch("referenceBaselinePt")), 14#
                equation.Paragraphs(1).Range.InsertParagraphAfter
            Next index
            doc.Range(doc.Content.End - 1, doc.Content.End - 1).Text = "BODY_AFTER"
            VTUpdateCurrentWordNumberFields doc
            Application.ScreenUpdating = True
            For Each position In Array(1, 25, 50)
                id = "91919191-9191-4191-8191-" & Format$(CLng(position), "000000000000")
                For sample = 1 To 10
                    numbered = sample Mod 2 = 0
                    stage = "apply-" & CStr(native) & "-" & CStr(chapter) & "-" & CStr(position) & "-" & CStr(sample)
                    Set dispatch = VTRegressionNumberingEditDispatch(doc, id, native, numbered)
                    If native And position = 1 And sample = 1 Then
                        VTDeleteWordMetadataPayload doc, id
                        VTDeleteWordLatexPayload doc, id
                        VTDeleteWordOmmlPayload doc, id
                        VTDeleteDocumentVariable doc, VTWordFormatVariableName(id)
                    End If
                    VTPrepareNativeRollbackDocument doc
                    VTWordPerformanceStart
                    VTWordPerformanceEnable dispatch
                    started = Timer
                    VTCommitWordDispatch CStr(dispatch("sessionId")), dispatch
                    elapsed = VTTimerElapsedSeconds(started, Timer) * 1000#
                    VTWordPerformanceFlush CStr(dispatch("sessionId"))
                    report = report & "sample|" & CStr(native) & "|" & CStr(chapter) & "|" & _
                        CStr(position) & "|" & CStr(numbered) & "|" & Format$(elapsed, "0.0") & vbLf
                    VTWriteTextAtomic root & "/progress.txt", stage & vbLf & report
                    VTRegressionAssertNumberingEdit doc, native, chapter, id, numbered
                Next sample
            Next position
            ' Both transitions must restore the old paragraph, number and edit
            ' payload if the Word transaction fails after removing the source.
            id = "91919191-9191-4191-8191-000000000001"
            For sample = 0 To 1
                stage = "rollback-" & CStr(native) & "-" & CStr(chapter) & "-" & CStr(sample)
                Set dispatch = VTRegressionNumberingEditDispatch(doc, id, native, sample = 1)
                On Error Resume Next
                VTCommitWordDispatch CStr(dispatch("sessionId")), dispatch, True
                errorNumber = Err.Number
                Err.Clear
                On Error GoTo Failed
                If errorNumber = 0 Then Err.Raise 5, , "Injected edit did not fail."
                VTRegressionAssertNumberingEdit doc, native, chapter, id, sample = 0
                If sample = 0 Then
                    Set dispatch = VTRegressionNumberingEditDispatch(doc, id, native, False)
                    VTCommitWordDispatch CStr(dispatch("sessionId")), dispatch
                End If
            Next sample
            stage = "save-reopen"
            doc.Save
            doc.Close SaveChanges:=wdDoNotSaveChanges
            Set doc = Documents.Open(FileName:=outputPath, ReadOnly:=False, AddToRecentFiles:=False)
            VTRegressionAssertNumberingEdit doc, native, chapter, id, False
            doc.Close SaveChanges:=wdDoNotSaveChanges
            Set doc = Nothing
        Next formatIndex
    Next representation
    VTWriteTextAtomic root & "/result.txt", "PASS" & vbLf & report
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = True
    VTWriteTextAtomic root & "/result.txt", "FAIL" & vbLf & "stage=" & stage & vbLf & _
        "error=" & CStr(errorNumber) & ": " & errorDescription & vbLf & report
End Sub

Private Function VTRegressionNumberingEditDispatch( _
    ByVal doc As Document, ByVal id As String, _
    ByVal native As Boolean, ByVal numbered As Boolean) As Object
    Dim dispatch As New Collection
    Dim row As Variant
    Dim splitAt As Long
    Dim sessionId As String
    Dim source As Range
    Dim root As String
    Dim contents As String
    Dim shape As InlineShape
    Dim parsedId As String
    Dim display As String
    Dim wasNumbered As Boolean
    root = VTApplicationSupportRoot() & "/Tests/numbering-edit"
    contents = VTReadText(root & "/" & id & "-" & IIf(numbered, "on", "off") & ".txt", 262144)
    For Each row In Split(Replace$(contents, vbCr, ""), vbLf)
        splitAt = InStr(1, CStr(row), "=", vbBinaryCompare)
        If splitAt > 1 Then dispatch.Add Mid$(CStr(row), splitAt + 1), Left$(CStr(row), splitAt - 1)
    Next row
    sessionId = VTNewUuidV4()
    dispatch.Add sessionId, "sessionId"
    dispatch.Add "edit", "mode"
    dispatch.Add IIf(native, "1", "0"), "nativeEquation"
    dispatch.Add "1", "numberingOnly"
    dispatch.Add "1", "performanceTrace"
    dispatch.Add VTWordDocumentIdentityForDocument(doc), "sourceDocumentId"
    If native Then
        If doc.OMaths.Count >= CLng(Right$(id, 12)) Then
            Set source = doc.OMaths(CLng(Right$(id, 12))).Range.Duplicate
            VTAddWordEditRangeBookmark source, sessionId
        End If
        dispatch.Add VTWordEditBookmarkName(sessionId), "sourceMarker"
    Else
        For Each shape In doc.InlineShapes
            If VTTryParseFormulaReference(shape.Title, parsedId, display, wasNumbered) Then
                If parsedId = id Then
                    VTAddWordEditRangeBookmark shape.Range, sessionId
                    Exit For
                End If
            End If
        Next shape
        dispatch.Add VTWordEditBookmarkName(sessionId), "sourceMarker"
    End If
    Set VTRegressionNumberingEditDispatch = dispatch
End Function

Private Sub VTRegressionAssertNumberingEdit( _
    ByVal doc As Document, ByVal native As Boolean, ByVal chapter As Boolean, _
    ByVal changedId As String, ByVal numbered As Boolean)
    Dim targets As Collection
    Dim target As Range
    Dim index As Long
    Dim expected As String
    Dim display As String
    Dim storedNumbered As Boolean
    Dim metadata As String
    Dim expectedCount As Long
    expectedCount = 49
    If numbered Then expectedCount = 50
    If native Then
        Set targets = VTCurrentNumberTargets(doc, VT_WORD_REFERENCE_KIND_NATIVE)
        If doc.OMaths.Count <> 50 Or doc.InlineShapes.Count <> 0 Then Err.Raise 5, , "Native equations were lost."
    Else
        Set targets = VTCurrentNumberTargets(doc, VT_WORD_REFERENCE_KIND_IMAGE)
        If doc.InlineShapes.Count <> 50 Or doc.OMaths.Count <> 0 Then Err.Raise 5, , "Image equations were lost."
    End If
    If targets.Count <> expectedCount Then Err.Raise 5, , "Wrong number inventory: " & CStr(targets.Count)
    For index = 1 To targets.Count
        expected = CStr(index)
        If chapter Then expected = "1." & expected
        Set target = targets(index)
        If Trim$(target.Text) <> expected Then Err.Raise 5, , "Stale number: " & target.Text & "; expected " & expected
    Next index
    If native Then
        Set target = doc.OMaths(CLng(Right$(changedId, 12))).Range.Duplicate
        storedNumbered = Not VTPureNativeNumberField(target) Is Nothing
        If storedNumbered <> numbered Then Err.Raise 5, , "Native number did not follow the checkbox."
        If doc.Bookmarks.Exists(VTNativeFormulaBookmarkName(changedId)) Or _
           VTTryReadWordMetadataPayload(doc, changedId, metadata) Then _
            Err.Raise 5, , "Private native formula state remains."
    Else
        If Not VTTryReadWordFormulaFormat(doc, changedId, display, storedNumbered) Then Err.Raise 5
        If storedNumbered <> numbered Then Err.Raise 5, , "Numbering payload did not follow the checkbox."
        If Not VTTryReadWordMetadataPayload(doc, changedId, metadata) Then Err.Raise 5, , "Edit metadata was deleted."
    End If
    If doc.Tables.Count <> 0 Or doc.Paragraphs.Count <> 52 Then Err.Raise 5, , "Paragraph topology changed."
    If Trim$(Replace$(doc.Paragraphs(1).Range.Text, vbCr, "")) <> "BODY_BEFORE" Or _
       Trim$(Replace$(doc.Paragraphs.Last.Range.Text, vbCr, "")) <> "BODY_AFTER" Then Err.Raise 5, , "Ordinary text changed."
End Sub

Public Sub VisualTeX_RunNumberedNativeComplexStructureRegression()
    Dim fixtureIds As Variant
    Dim fixtureNames As Variant
    Dim doc As Document
    Dim insertion As Range
    Dim equation As Range
    Dim formulaBody As Range
    Dim placeholder As Range
    Dim numberField As Field
    Dim fixturePath As String
    Dim savedPath As String
    Dim formulaId As String
    Dim fixtureRoot As String
    Dim stage As String
    Dim report As String
    Dim index As Long
    Dim errorNumber As Long
    Dim errorDescription As String

    fixtureIds = Array( _
        "22222222-2222-4222-8222-222222222222", _
        "23232323-2323-4323-8323-232323232323", _
        "24242424-2424-4424-8424-242424242424", _
        "25252525-2525-4525-8525-252525252525", _
        "26262626-2626-4626-8626-262626262626", _
        "27272727-2727-4727-8727-272727272727")
    fixtureNames = Array( _
        "binomial", "nested-delimiters", "matrix", "cases", _
        "aligned-array", "nary-fraction-root")
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    savedPath = ThisDocument.Path & _
        "/numbered-native-complex-structures.docx"

    On Error GoTo Failed
    VTBeginWordInternalMutation
    Set doc = Documents.Add(Visible:=True)
    doc.ActiveWindow.View.Type = wdPrintView
    doc.Activate
    VTSetDocumentVariable _
        doc, VT_WORD_NUMBERING_MODE_VARIABLE, VT_WORD_NUMBERING_MODE_SEQUENCE

    For index = LBound(fixtureIds) To UBound(fixtureIds)
        stage = "insert-" & CStr(fixtureNames(index))
        fixturePath = VTApplicationSupportRoot() & _
            "/NativeDocuments/" & CStr(fixtureIds(index)) & ".docx"
        If Not VTPathFileExists(fixturePath) Then
            Err.Raise vbObjectError + 7607, "VisualTeX", _
                "The complex numbered OMML fixture is missing: " & _
                CStr(fixtureNames(index))
        End If
        Set insertion = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        Set equation = VTInsertCachedNativeEquationFast( _
            insertion, fixturePath, 14#)
        If equation Is Nothing Or equation.OMaths.Count <> 1 Then
            Err.Raise vbObjectError + 7607, "VisualTeX", _
                "Word did not import the complex numbered OMML fixture: " & _
                CStr(fixtureNames(index))
        End If
        Set placeholder = VTNativeEquationArrayPlaceholderRange(equation)
        If placeholder Is Nothing Then
            Err.Raise vbObjectError + 7607, "VisualTeX", _
                "The owned number slot was not resolved for: " & _
                CStr(fixtureNames(index))
        End If
        formulaId = VTNewUuidV4()
        Set equation = VTWritePureNativeNumber(equation, formulaId)
        VTFinishPureNativeEquation equation, formulaId
        Set formulaBody = VTNativeEquationFormulaContentRange(equation)
        If formulaBody Is Nothing Then
            Err.Raise vbObjectError + 7607, "VisualTeX", _
                "The complex OMML body was not resolved for: " & _
                CStr(fixtureNames(index))
        End If
        equation.Font.Size = CSng(14 + index)
        equation.Paragraphs(1).Range.InsertParagraphAfter
        report = report & CStr(fixtureNames(index)) & "=PASS" & vbLf
    Next index

    stage = "verify-sequence"
    VTUpdateCurrentWordNumberFields doc
    If doc.OMaths.Count <> UBound(fixtureIds) - LBound(fixtureIds) + 1 Then
        Err.Raise vbObjectError + 7607, "VisualTeX", _
            "Complex numbered OMML insertion lost or duplicated equations."
    End If
    For index = 1 To doc.OMaths.Count
        Set numberField = VTPureNativeNumberField(doc.OMaths(index).Range)
        If numberField Is Nothing Or _
           Trim$(numberField.Result.Text) <> CStr(index) Then
            Err.Raise vbObjectError + 7607, "VisualTeX", _
                "Complex numbered OMML is not numbered in document order at " & _
                CStr(index) & "."
        End If
    Next index

    VTRegressionAssertNativeBodyFontSizes doc, _
        "complex-sequence-fonts", 6, 14!, 1!
    VTRegressionAssertNativeNumberFontSizes doc, _
        "complex-sequence-number-fonts", 6, 14!, 1!
    stage = "complex-chapter-format-fonts"
    VTSetEquationNumberingFormat _
        doc, VT_WORD_NUMBERING_MODE_CHAPTER, "-"
    VTUpdateEquationNumbersCore doc, False
    VTRegressionAssertNativeBodyFontSizes doc, stage, 6, 14!, 1!
    VTRegressionAssertNativeNumberFontSizes doc, stage, 6, 14!, 1!
    stage = "complex-section-format-fonts"
    VTSetEquationNumberingFormat _
        doc, VT_WORD_NUMBERING_MODE_SECTION, "."
    VTUpdateEquationNumbersCore doc, False
    VTRegressionAssertNativeBodyFontSizes doc, stage, 6, 14!, 1!
    VTRegressionAssertNativeNumberFontSizes doc, stage, 6, 14!, 1!
    stage = "complex-sequence-reset-fonts"
    VTSetEquationNumberingFormat _
        doc, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    VTUpdateEquationNumbersCore doc, False
    VTRegressionAssertNativeBodyFontSizes doc, stage, 6, 14!, 1!
    VTRegressionAssertNativeNumberFontSizes doc, stage, 6, 14!, 1!
    report = report & "format-font-preservation=PASS" & vbLf

    stage = "save-reopen-f9"
    doc.SaveAs2 FileName:=savedPath, FileFormat:=wdFormatXMLDocument, _
        AddToRecentFiles:=False
    doc.Close SaveChanges:=wdDoNotSaveChanges
    Set doc = Documents.Open( _
        FileName:=savedPath, AddToRecentFiles:=False, Visible:=True)
    doc.Fields.Update
    VTUpdateCurrentWordNumberFields doc
    If doc.OMaths.Count <> UBound(fixtureIds) - LBound(fixtureIds) + 1 Then
        Err.Raise vbObjectError + 7607, "VisualTeX", _
            "Save and reopen changed the complex numbered OMML count."
    End If
    For index = 1 To doc.OMaths.Count
        Set numberField = VTPureNativeNumberField(doc.OMaths(index).Range)
        If numberField Is Nothing Or _
           Trim$(numberField.Result.Text) <> CStr(index) Then
            Err.Raise vbObjectError + 7607, "VisualTeX", _
                "F9 changed a complex native Equation number at " & _
                CStr(index) & "."
        End If
    Next index
    VTRegressionAssertNativeBodyFontSizes doc, _
        "complex-save-reopen-fonts", 6, 14!, 1!
    VTRegressionAssertNativeNumberFontSizes doc, _
        "complex-save-reopen-number-fonts", 6, 14!, 1!
    report = report & "save-reopen-f9=PASS" & vbLf & "status=PASS" & vbLf
    GoTo Finished

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    report = report & "status=FAIL" & vbLf & _
        "stage=" & stage & vbLf & _
        "error=" & CStr(errorNumber) & ":" & errorDescription & vbLf
    On Error Resume Next
    If Not doc Is Nothing Then
        VTWriteTextAtomic _
            fixtureRoot & "/numbered-native-complex-structures-failure.xml", _
            doc.Content.WordOpenXML
    End If
Finished:
    On Error Resume Next
    VTWriteTextAtomic _
        fixtureRoot & "/numbered-native-complex-structures-result.txt", report
    If Not doc Is Nothing Then doc.Close SaveChanges:=wdDoNotSaveChanges
    VTEndWordInternalMutation
    On Error GoTo 0
End Sub

Private Sub VTRegressionAssertNativeNumberFontSizes( _
    ByVal documentObject As Document, ByVal assertionStage As String, _
    Optional ByVal expectedCount As Long = 3, _
    Optional ByVal firstExpectedSize As Single = 15!, _
    Optional ByVal expectedSizeStep As Single = 2!)

    Dim formulaIndex As Long
    Dim numberField As Field
    Dim visibleNumber As Range
    Dim expectedSize As Single
    Dim actualSize As Single

    If documentObject Is Nothing Or _
       documentObject.OMaths.Count < expectedCount Then
        Err.Raise vbObjectError + 7607, "VisualTeX regression", _
            assertionStage & ": missing native Equation fixtures."
    End If
    For formulaIndex = 1 To expectedCount
        expectedSize = firstExpectedSize + _
            CSng(formulaIndex - 1) * expectedSizeStep
        Set numberField = VTPureNativeNumberField( _
            documentObject.OMaths(formulaIndex).Range)
        Set visibleNumber = VTPureNativeVisibleNumberRange( _
            documentObject.OMaths(formulaIndex).Range, numberField)
        If visibleNumber Is Nothing Then
            Err.Raise vbObjectError + 7607, "VisualTeX regression", _
                assertionStage & ": missing native Equation number at " & _
                CStr(formulaIndex) & "."
        End If
        actualSize = visibleNumber.Font.Size
        If Abs(actualSize - expectedSize) > 0.1 Then
            Err.Raise vbObjectError + 7607, "VisualTeX regression", _
                assertionStage & ": native Equation number font size changed" & _
                " [index=" & CStr(formulaIndex) & _
                "; actual=" & CStr(actualSize) & _
                "; expected=" & CStr(expectedSize) & "]."
        End If
    Next formulaIndex
End Sub

Private Sub VTRegressionAssertNativeBodyFontSizes( _
    ByVal documentObject As Document, ByVal assertionStage As String, _
    Optional ByVal expectedCount As Long = 3, _
    Optional ByVal firstExpectedSize As Single = 15!, _
    Optional ByVal expectedSizeStep As Single = 2!)

    Dim formulaIndex As Long
    Dim formulaBody As Range
    Dim expectedSize As Single
    Dim actualSize As Single

    If documentObject Is Nothing Or _
       documentObject.OMaths.Count < expectedCount Then
        Err.Raise vbObjectError + 7607, "VisualTeX regression", _
            assertionStage & ": missing native Equation fixtures."
    End If
    For formulaIndex = 1 To expectedCount
        expectedSize = firstExpectedSize + _
            CSng(formulaIndex - 1) * expectedSizeStep
        Set formulaBody = VTNativeEquationFormulaContentRange( _
            documentObject.OMaths(formulaIndex).Range)
        If formulaBody Is Nothing Then
            Err.Raise vbObjectError + 7607, "VisualTeX regression", _
                assertionStage & ": missing native Equation body at " & _
                CStr(formulaIndex) & "."
        End If
        actualSize = formulaBody.Font.Size
        If Abs(actualSize - expectedSize) > 0.1 Then
            Err.Raise vbObjectError + 7607, "VisualTeX regression", _
                assertionStage & ": native Equation body font size changed" & _
                " [index=" & CStr(formulaIndex) & _
                "; actual=" & CStr(actualSize) & _
                "; expected=" & CStr(expectedSize) & "]."
        End If
    Next formulaIndex
End Sub

Public Sub VisualTeX_RunNativeFormatContractRegression()
    Dim doc As Document
    Dim insertion As Range
    Dim math As OMath
    Dim shape As InlineShape
    Dim secondShape As InlineShape
    Dim placeRef As Field
    Dim imageSequence As Field
    Dim numberRange As Range
    Dim numberTail As Range
    Dim placeholder As Range
    Dim marker As Range
    Dim source As Range
    Dim id As String
    Dim secondId As String
    Dim index As Long
    Dim candidate As Bookmark
    Dim numberCreated As Boolean
    Dim numberBookmarks As Collection
    Dim savedPath As String
    Dim anchor As Long
    Dim imageReference As Range
    Dim nativeTargets As Collection
    Dim imageTargets As Collection
    Dim headingRange As Range
    Dim emptyRibbonControl As IRibbonControl
    Dim referenceMenuXml As Variant
    Dim report As String
    Dim stage As String
    Dim errorNumber As Long
    Dim errorDescription As String
    ' The candidate template is tested alongside the currently installed add-in.
    ' Suspend that add-in's document-open migration for these disposable fixtures.
    On Error Resume Next
    Application.Run "VisualTeX.dotm!VisualTeX_DisableWordEventsForRegression"
    On Error GoTo Failed
    VTBeginWordInternalMutation
    Set doc = Documents.Add(Visible:=True)
    VTSetDocumentVariable doc, VT_WORD_NUMBERING_MODE_VARIABLE, VT_WORD_NUMBERING_MODE_SEQUENCE
    stage = "existing-number-cache-shell"
    id = VTNewUuidV4()
    Set insertion = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    insertion.Text = "c=0"
    Set insertion = doc.OMaths.Add(insertion)
    Set math = insertion.OMaths(1)
    math.BuildUp
    Set numberRange = VTImportNativeNumberShell(math.Range, id)
    Set marker = VTNativeEquationArrayMarkerRange(numberRange)
    If marker Is Nothing Then Err.Raise vbObjectError + 7607
    Set placeholder = doc.Range(marker.End, numberRange.End)
    With placeholder.Find
        .ClearFormatting
        .Text = "0"
        .Forward = False
        .Wrap = wdFindStop
        .Format = False
    End With
    If Not placeholder.Find.Execute Then Err.Raise vbObjectError + 7607
    placeholder.Text = VT_WORD_NATIVE_NUMBER_CACHE_PLACEHOLDER
    Set numberRange = VTNativeMathRoot(placeholder.Paragraphs(1).Range.OMaths(1)).Range.Duplicate
    Set placeholder = VTNativeEquationArrayPlaceholderRange(numberRange)
    If placeholder Is Nothing Then
        Err.Raise vbObjectError + 7607, , _
            "The numbered native cache did not expose its owned placeholder."
    End If
    Set numberRange = VTWritePureNativeNumber(numberRange, id)
    If VTPureNativeNumberField(numberRange) Is Nothing Then
        Err.Raise vbObjectError + 7607, , _
            "The numbered native cache placeholder was not replaced by SEQ."
    End If
    VTFinishPureNativeEquation numberRange, id
    numberRange.Paragraphs(1).Range.Delete
    report = "existing-number-cache-shell=PASS" & vbLf
    stage = "insert-pure-native"
    For index = 1 To 3
        id = VTNewUuidV4()
        Set insertion = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        insertion.Text = "x=" & CStr(index)
        Set insertion = doc.OMaths.Add(insertion)
        Set math = insertion.OMaths(1)
        math.BuildUp
        Set numberRange = VTWritePureNativeNumber(math.Range, id)
        VTFinishPureNativeEquation numberRange, id
        numberRange.Font.Size = CSng(13 + index * 2)
        Set source = VTNativeEquationFormulaContentRange(numberRange)
        If source Is Nothing Then Err.Raise vbObjectError + 7607
        numberRange.Paragraphs(1).Range.InsertParagraphAfter
    Next index
    stage = "verify-pure-native"
    VTUpdateCurrentWordNumberFields doc
    If doc.OMaths.Count <> 3 Or doc.Frames.Count <> 0 Or doc.Tables.Count <> 0 Then
        Err.Raise vbObjectError + 7607, , "Native insertion created a non-native host."
    End If
    For index = 1 To doc.OMaths.Count
        Set math = doc.OMaths(index)
        If VTPureNativeNumberField(math.Range) Is Nothing Then Err.Raise vbObjectError + 7607
        If Trim$(VTPureNativeNumberField(math.Range).Result.Text) <> CStr(index) Then
            Err.Raise vbObjectError + 7607, , "Native sequence is not in document order."
        End If
    Next index
    For Each candidate In doc.Bookmarks
        If Left$(candidate.Name, 3) = "VT_" Then Err.Raise vbObjectError + 7607
    Next candidate
    If doc.Variables.Count <> 1 Then Err.Raise vbObjectError + 7607, , "Native metadata survived."
    VTRegressionAssertNativeBodyFontSizes doc, "initial-native-fonts"
    VTRegressionAssertNativeNumberFontSizes doc, "initial-number-fonts"
    report = report & "pure-native=PASS" & vbLf
    stage = "middle-insertion-and-deletion"
    doc.Paragraphs(1).Range.InsertParagraphAfter
    anchor = doc.Paragraphs(2).Range.Start
    Set insertion = doc.Range(anchor, anchor)
    insertion.Text = "y"
    Set insertion = doc.Range(anchor, anchor + 1)
    Set insertion = doc.OMaths.Add(insertion)
    id = VTNewUuidV4()
    Set numberRange = VTWritePureNativeNumber(insertion.OMaths(1).Range, id)
    VTFinishPureNativeEquation numberRange, id
    VTUpdateCurrentWordNumberFields doc
    If Trim$(VTPureNativeNumberField(doc.OMaths(4).Range).Result.Text) <> "4" Then
        Err.Raise vbObjectError + 7607, , "Middle insertion did not renumber native math."
    End If
    numberRange.Paragraphs(1).Range.Delete
    VTUpdateCurrentWordNumberFields doc
    If Trim$(VTPureNativeNumberField(doc.OMaths(3).Range).Result.Text) <> "3" Then
        Err.Raise vbObjectError + 7607, , "Deletion did not renumber native math."
    End If
    report = report & "middle-insert-delete=PASS" & vbLf
    stage = "insert-image-paragraph"
    id = VTNewUuidV4()
    Set insertion = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    Set shape = doc.InlineShapes.AddPicture(FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, Range:=insertion)
    shape.Width = 80!
    shape.Height = 20!
    shape.Title = VTFormulaReference(id, "block", True)
    shape.AlternativeText = VT_METADATA_PREFIX & "e30"
    Set numberRange = VTWriteSingleParagraphImageNumber(shape, id)
    VTUpdateCurrentWordNumberFields doc
    If doc.Frames.Count <> 0 Then Err.Raise vbObjectError + 7607, , "Image created a Frame."
    Set placeRef = VTImagePlaceRef(shape.Range)
    If placeRef Is Nothing Then Err.Raise vbObjectError + 7607
    Set imageSequence = VTImageSequenceField(placeRef, False)
    If imageSequence Is Nothing Or _
       Trim$(VTCurrentImageNumberRange(placeRef).Text) <> "1" Then
        Err.Raise vbObjectError + 7607, , _
            "The image sequence did not start independently at 1: " & _
            VTCurrentImageNumberRange(placeRef).Text
    End If
    If Right$(placeRef.Code.Text, 1) <> ")" Then
        Err.Raise vbObjectError + 7607, , _
            "VisualTeXPlaceRef contains trailing whitespace."
    End If
    Set numberTail = doc.Range( _
        VTEquationFieldEnd(placeRef), shape.Range.Paragraphs(1).Range.End - 1)
    If Len(numberTail.Text) <> 0 Then
        Err.Raise vbObjectError + 7607, , _
            "The image paragraph contains text after VisualTeXPlaceRef."
    End If
    report = report & "image-paragraph=PASS" & vbLf
    stage = "independent-image-sequence"
    shape.Range.Paragraphs(1).Range.InsertParagraphAfter
    secondId = VTNewUuidV4()
    Set insertion = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    Set secondShape = doc.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), LinkToFile:=False, _
        SaveWithDocument:=True, Range:=insertion)
    secondShape.Width = 80!
    secondShape.Height = 20!
    secondShape.Title = VTFormulaReference(secondId, "block", True)
    secondShape.AlternativeText = VT_METADATA_PREFIX & "e30"
    Set numberRange = VTWriteSingleParagraphImageNumber(secondShape, secondId)
    VTUpdateCurrentWordNumberFields doc
    If Trim$(VTCurrentImageNumberRange( _
       VTImagePlaceRef(secondShape.Range)).Text) <> "2" Then
        Err.Raise vbObjectError + 7607, , _
            "The second image did not advance its private sequence to 2."
    End If
    For index = 1 To 3
        Set math = doc.OMaths(index)
        If Trim$(VTPureNativeNumberField(math.Range).Result.Text) <> _
           CStr(index) Then
            Err.Raise vbObjectError + 7607, , _
                "An image advanced the native OMML Equation sequence."
        End If
    Next index
    secondShape.Range.Paragraphs(1).Range.Delete
    VTUpdateCurrentWordNumberFields doc
    Set shape = doc.InlineShapes(1)
    If Trim$(VTCurrentImageNumberRange( _
       VTImagePlaceRef(shape.Range)).Text) <> "1" Then
        Err.Raise vbObjectError + 7607, , _
            "Deleting the second image did not compact only the image sequence."
    End If
    report = report & "independent-image-omml-sequences=PASS" & vbLf
    stage = "format-no-heading-chapter-dash"
    VTSetEquationNumberingFormat _
        doc, VT_WORD_NUMBERING_MODE_CHAPTER, "-"
    VTUpdateEquationNumbersCore doc, False
    VTRegressionAssertNativeBodyFontSizes doc, _
        "chapter-dash-format-fonts"
    VTRegressionAssertNativeNumberFontSizes doc, _
        "chapter-dash-format-number-fonts"
    Set nativeTargets = VTCurrentNumberTargets( _
        doc, VT_WORD_REFERENCE_KIND_NATIVE)
    Set imageTargets = VTCurrentNumberTargets( _
        doc, VT_WORD_REFERENCE_KIND_IMAGE)
    If nativeTargets.Count <> 3 Or imageTargets.Count <> 1 Then
        Err.Raise vbObjectError + 7607, , _
            "The typed Equation inventories changed during format migration."
    End If
    If Trim$(nativeTargets(1).Text) <> "0" & ChrW(8208) & "1" Or _
       Trim$(imageTargets(1).Text) <> "0" & ChrW(8208) & "1" Then
        Err.Raise vbObjectError + 7607, , _
            "Chapter-dash format silently fell back to sequence before Heading 1."
    End If
    stage = "format-no-heading-section-dot"
    VTSetEquationNumberingFormat _
        doc, VT_WORD_NUMBERING_MODE_SECTION, "."
    VTUpdateEquationNumbersCore doc, False
    VTRegressionAssertNativeBodyFontSizes doc, _
        "section-dot-format-fonts"
    VTRegressionAssertNativeNumberFontSizes doc, _
        "section-dot-format-number-fonts"
    Set nativeTargets = VTCurrentNumberTargets( _
        doc, VT_WORD_REFERENCE_KIND_NATIVE)
    Set imageTargets = VTCurrentNumberTargets( _
        doc, VT_WORD_REFERENCE_KIND_IMAGE)
    If Trim$(nativeTargets(1).Text) <> "0.0.1" Or _
       Trim$(imageTargets(1).Text) <> "0.0.1" Then
        Err.Raise vbObjectError + 7607, , _
            "Section-dot format silently fell back to sequence before headings."
    End If
    stage = "format-set-sequence"
    VTSetEquationNumberingFormat _
        doc, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    stage = "format-rebuild-sequence"
    VTUpdateEquationNumbersCore doc, False
    VTRegressionAssertNativeBodyFontSizes doc, _
        "sequence-format-fonts"
    VTRegressionAssertNativeNumberFontSizes doc, _
        "sequence-format-number-fonts"
    Set nativeTargets = VTCurrentNumberTargets( _
        doc, VT_WORD_REFERENCE_KIND_NATIVE)
    Set imageTargets = VTCurrentNumberTargets( _
        doc, VT_WORD_REFERENCE_KIND_IMAGE)
    If Trim$(nativeTargets(1).Text) <> "1" Or _
       Trim$(imageTargets(1).Text) <> "1" Then
        Err.Raise vbObjectError + 7607, , _
            "Sequence format did not remove the chapter prefix."
    End If
    stage = "format-insert-heading"
    doc.Paragraphs(1).Range.InsertParagraphBefore
    Set insertion = doc.Paragraphs(1).Range.Duplicate
    insertion.End = insertion.End - 1
    insertion.Text = "VisualTeX chapter"
    Set headingRange = doc.Paragraphs(1).Range.Duplicate
    headingRange.Style = wdStyleHeading1
    stage = "format-heading-chapter-dash"
    VTSetEquationNumberingFormat _
        doc, VT_WORD_NUMBERING_MODE_CHAPTER, "-"
    VTUpdateEquationNumbersCore doc, False
    VTRegressionAssertNativeBodyFontSizes doc, _
        "heading-chapter-format-fonts"
    VTRegressionAssertNativeNumberFontSizes doc, _
        "heading-chapter-format-number-fonts"
    Set nativeTargets = VTCurrentNumberTargets( _
        doc, VT_WORD_REFERENCE_KIND_NATIVE)
    Set imageTargets = VTCurrentNumberTargets( _
        doc, VT_WORD_REFERENCE_KIND_IMAGE)
    If Trim$(nativeTargets(1).Text) <> "1" & ChrW(8208) & "1" Or _
       Trim$(imageTargets(1).Text) <> "1" & ChrW(8208) & "1" Then
        Err.Raise vbObjectError + 7607, , _
            "An unnumbered Heading 1 did not create the implicit chapter prefix."
    End If
    stage = "format-heading-reset-sequence"
    VTSetEquationNumberingFormat _
        doc, VT_WORD_NUMBERING_MODE_SEQUENCE, "."
    VTUpdateEquationNumbersCore doc, False
    VTRegressionAssertNativeBodyFontSizes doc, _
        "heading-sequence-format-fonts"
    VTRegressionAssertNativeNumberFontSizes doc, _
        "heading-sequence-format-number-fonts"
    stage = "typed-reference-menu"
    VTWordRibbonCrossReferenceMenuContent _
        emptyRibbonControl, referenceMenuXml
    If InStr(1, CStr(referenceMenuXml), _
       "CrossReference.Omml", vbBinaryCompare) = 0 Or _
       InStr(1, CStr(referenceMenuXml), _
       "CrossReference.Image", vbBinaryCompare) = 0 Or _
       InStr(1, CStr(referenceMenuXml), _
       "tag=""omml|1""", vbBinaryCompare) = 0 Or _
       InStr(1, CStr(referenceMenuXml), _
       "tag=""image|1""", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7607, , _
            "The Ribbon Equation-reference menu did not separate OMML and image formulas."
    End If
    doc.Paragraphs(1).Range.Delete
    Set shape = doc.InlineShapes(1)
    report = report & "numbering-format-migration=PASS" & vbLf & _
        "typed-reference-inventory=PASS" & vbLf & _
        "typed-reference-menu=PASS" & vbLf
    stage = "reference-and-f9"
    shape.Range.Paragraphs(1).Range.InsertParagraphAfter
    Set insertion = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    Set source = VTInsertEquationNumberReferenceAtRange( _
        insertion, 2, True, VT_WORD_REFERENCE_KIND_NATIVE)
    doc.Fields.Update
    VTUpdateCurrentWordNumberFields doc
    If source.Text <> "(2)" Then Err.Raise vbObjectError + 7607, , "Native body reference changed."
    report = report & "reference-f9=PASS" & vbLf
    source.Paragraphs(1).Range.InsertParagraphAfter
    Set insertion = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    Set imageReference = VTInsertEquationNumberReferenceAtRange( _
        insertion, 1, True, VT_WORD_REFERENCE_KIND_IMAGE)
    If imageReference.Text <> "(1)" Then Err.Raise vbObjectError + 7607, , "Image REF is not result-only."
    report = report & "image-reference=PASS" & vbLf
    VTDeleteDocumentVariable doc, VT_WORD_NUMBERING_MODE_VARIABLE
    VTDeleteDocumentVariable doc, VT_WORD_NUMBERING_SEPARATOR_VARIABLE
    stage = "save"
    doc.SaveAs2 FileName:=ThisDocument.Path & "/native-format-contract.docx", _
        FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
    savedPath = doc.FullName
    doc.Close SaveChanges:=wdDoNotSaveChanges
    Set doc = Documents.Open(FileName:=savedPath, AddToRecentFiles:=False)
    doc.Fields.Update
    VTUpdateCurrentWordNumberFields doc
    Set shape = doc.InlineShapes(1)
    If doc.Frames.Count <> 0 Or doc.Paragraphs.Count <> 6 Then
        Err.Raise vbObjectError + 7607, , "Reopen changed paragraph topology."
    End If
    If VTCurrentImageNumberRange(VTImagePlaceRef(shape.Range)).Text <> "1" Then Err.Raise vbObjectError + 7607
    report = report & "save-reopen-f9=PASS" & vbLf
    stage = "conversion-parse-image"
    If Not VTTryParseFormulaReference(shape.Title, id, savedPath, numberCreated) Then Err.Raise vbObjectError + 7607
    stage = "conversion-capture-bookmarks"
    Set numberBookmarks = VTCaptureImageNumberBookmarks(shape, id)
    report = report & "captured-number-bookmarks=" & CStr(numberBookmarks.Count) & vbLf
    stage = "conversion-prepare-image-carrier"
    Set shape = VTPrepareNumberedImageCarrierForNativeFastConversion(doc, shape, id)
    anchor = shape.Range.Start
    shape.Delete
    stage = "conversion-create-native"
    Set insertion = doc.Range(anchor, anchor)
    insertion.Text = "z=4"
    Set insertion = doc.OMaths.Add(insertion)
    stage = "conversion-write-native-number"
    Set numberRange = VTWritePureNativeNumber(insertion.OMaths(1).Range, id)
    VTWriteTextAtomic VTApplicationSupportRoot() & "/Tests/converted-native.xml", numberRange.WordOpenXML
    report = report & "math-count=" & CStr(numberRange.OMaths.Count) & ";fields=" & CStr(numberRange.Fields.Count) & vbLf
    stage = "conversion-restore-bookmarks"
    VTRestoreNativeNumberBookmarks numberRange, numberBookmarks
    report = report & "restored-native-field=" & CStr(Not VTPureNativeNumberField(numberRange) Is Nothing) & vbLf
    stage = "conversion-finish-native"
    VTFinishPureNativeEquation numberRange, id
    stage = "conversion-update-fields"
    VTUpdateCurrentWordNumberFields doc
    stage = "conversion-assert-reference-text"
    Set source = doc.Paragraphs.Last.Range.Duplicate
    If InStr(1, source.Text, "(4)", vbBinaryCompare) = 0 Then _
        Err.Raise vbObjectError + 7607, , "Converted number broke its REF."
    stage = "conversion-assert-bookmarks"
    For Each candidate In doc.Bookmarks
        If Left$(candidate.Name, 3) = "VT_" Then Err.Raise vbObjectError + 7607, , "Converted native retained identity."
    Next candidate
    stage = "conversion-assert-no-images"
    If doc.InlineShapes.Count <> 0 Then Err.Raise vbObjectError + 7607
    stage = "conversion-assert-number-targets"
    Set nativeTargets = VTCurrentNumberTargets( _
        doc, VT_WORD_REFERENCE_KIND_NATIVE)
    If nativeTargets.Count <> 4 Then Err.Raise vbObjectError + 7607
    report = report & "converted-number-reference=PASS" & vbLf
    stage = "ordinary-native"
    doc.Paragraphs.Last.Range.InsertParagraphAfter
    Set insertion = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    insertion.Text = "a+b"
    Set insertion = doc.OMaths.Add(insertion)
    id = VTNewUuidV4()
    VTSetNativeFormulaBookmark doc, insertion, id
    VTSetWordFormulaFormat doc, id, "inline", False
    VTFinishPureNativeEquation insertion, id
    If insertion.OMaths.Count <> 1 Then Err.Raise vbObjectError + 7607
    doc.SaveAs2 FileName:=ThisDocument.Path & "/native-format-pure.docx", _
        FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
    report = report & "ordinary-native=PASS" & vbLf & "status=PASS" & vbLf
Finished:
    On Error Resume Next
    VTEndWordInternalMutation
    Application.Run "VisualTeX.dotm!VTInitializeWordEvents"
    VTWriteTextAtomic VTApplicationSupportRoot() & "/Tests/native-format-contract.txt", _
        report & "stage=" & stage & vbLf & "error=" & CStr(errorNumber) & _
        ":" & errorDescription
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Resume Finished
End Sub

Private Function VTNativeParityFirstGlyph(ByVal exact As Range) As Range
    Dim character As Range
    Dim value As String
    For Each character In exact.Characters
        value = character.Text
        If value <> vbCr And value <> vbLf And value <> vbTab And _
           value <> " " And value <> Chr$(7) Then
            Set VTNativeParityFirstGlyph = character.Duplicate
            Exit Function
        End If
    Next character
    Err.Raise vbObjectError + 7607, , "The formula has no visible glyph."
End Function

Private Function VTNativeParityScratchRoot() As String
    Dim root As String
    root = VTApplicationSupportRoot()
    VTNativeParityScratchRoot = Left$(root, InStr(1, root, "/Library/", vbBinaryCompare) - 1) & _
        "/Library/Group Containers/UBF8T346G9.Office/VisualTeX/Scratch"
End Function

Public Sub VisualTeX_RunNativeParityRegression()
    Dim doc As Document
    Dim original As Range
    Dim insertion As Range
    Dim inserted As Range
    Dim math As OMath
    Dim mark As Bookmark
    Dim variable As Variable
    Dim id As String
    Dim fixture As String
    Dim payload As String
    Dim originalText As String
    Dim stage As String
    Dim report As String
    Dim index As Long
    Dim originalColor As Long
    Dim originalSize As Single
    Dim originalStyle As String
    Dim xmlBodyStart As Long
    Dim xmlBodyEnd As Long
    Dim formulaBody As Range
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    VT_NATIVE_PARITY_DIAGNOSTIC = True
    VTWriteTextAtomic VTApplicationSupportRoot() & "/Tests/native-parity-debug.txt", "r108-diagnostic" & vbLf
    VTBeginWordInternalMutation
    fixture = VTApplicationSupportRoot() & "/Tests/numbering-edit/off.docx"
    payload = VTBase64UrlEncodeUtf8( _
        "<m:oMath xmlns:m=""http:" & "//schemas.openxmlformats.org/officeDocument/2006/math""><m:r><m:t>x</m:t></m:r></m:oMath>")
    Set doc = Documents.Add(Visible:=True)
    doc.Activate
    doc.Styles(wdStyleCaption).Font.Color = RGB(79, 129, 189)
    doc.Styles(wdStyleCaption).Font.Bold = True
    doc.Styles(wdStyleCaption).Font.Size = 9!
    stage = "native-first"
    Set original = doc.Range(0, 0)
    original.Text = "x^2+1"
    Set original = doc.Range(0, Len("x^2+1"))
    Set original = doc.OMaths.Add(original)
    original.OMaths(1).BuildUp
    original.OMaths(1).Type = wdOMathInline
    Set original = VTNativeMathRoot(doc.OMaths(1)).Range.Duplicate
    originalText = VTXmlTextContent(VTNativeMathXmlBody( _
        original.WordOpenXML, xmlBodyStart, xmlBodyEnd))
    For index = 1 To 3
        stage = "adjacent-insert-" & CStr(index)
        Set math = VTNativeMathRoot(doc.OMaths(doc.OMaths.Count))
        Set insertion = doc.Range(math.Range.End, math.Range.End)
        VTNativeParityTrace "test-before-insert-" & CStr(index), doc, insertion
        Set inserted = VTInsertNativeEquationAtRange( _
            insertion, payload, fixture, "inline", False, False)
        VTNativeParityTrace "test-after-insert-" & CStr(index), doc, inserted
        stage = "adjacent-verify-" & CStr(index)
        If doc.OMaths.Count <> 1 Then _
            Err.Raise vbObjectError + 7607, , "Adjacent equations differ from Word's native math zone."
        If Left$(VTXmlTextContent(VTNativeMathXmlBody( _
            VTNativeMathRoot(doc.OMaths(1)).Range.WordOpenXML, xmlBodyStart, xmlBodyEnd)), _
            Len(originalText)) <> originalText Then _
            Err.Raise vbObjectError + 7607, , "The original native equation changed."
        id = VTNewUuidV4()
        VTSetNativeFormulaBookmark doc, inserted, id
        VTSetWordLatexPayload doc, id, VTBase64UrlEncodeUtf8("x")
        VTSetWordFormulaFormat doc, id, "inline", False
        stage = "adjacent-cleanup-" & CStr(index)
        VTFinishPureNativeEquation inserted, id
        stage = "adjacent-caret-" & CStr(index)
        VTPlaceCaretAfterInlineNativeEquation inserted
    Next index
    If InStr(doc.Content.Text, ChrW(8288)) > 0 Then _
        Err.Raise vbObjectError + 7607, , "A native equation left an invisible anchor."
    report = report & "adjacent-native=PASS" & vbLf
    stage = "body-text"
    Selection.TypeText Text:="tail"
    If InStr(doc.OMaths(doc.OMaths.Count).Range.Text, "tail") > 0 Then _
        Err.Raise vbObjectError + 7607, , "Ordinary text entered OMath."
    report = report & "ordinary-text=PASS" & vbLf
    stage = "number-colour"
    doc.Content.InsertParagraphAfter
    Set insertion = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    Set inserted = VTInsertNativeEquationAtRange( _
        insertion, payload, fixture, "block", True, False)
    inserted.Font.Size = 15!
    inserted.Font.Color = wdColorAutomatic
    Set formulaBody = VTNativeParityFirstGlyph(inserted)
    originalColor = formulaBody.Font.Color
    originalSize = formulaBody.Font.Size
    originalStyle = inserted.Paragraphs(1).Range.Style.NameLocal
    id = VTNewUuidV4()
    Set inserted = VTWritePureNativeNumber(inserted, id)
    Set formulaBody = VTNativeParityFirstGlyph(VTNativeEquationFormulaContentRange(inserted))
    If formulaBody.Font.Color <> originalColor Or _
       formulaBody.Font.Size <> originalSize Or _
       inserted.Paragraphs(1).Range.Style.NameLocal <> originalStyle Then _
        Err.Raise vbObjectError + 7607, , "Numbering changed native formatting: colour=" & CStr(originalColor) & "/" & CStr(formulaBody.Font.Color) & "; size=" & CStr(originalSize) & "/" & CStr(formulaBody.Font.Size) & "; style=" & originalStyle & "/" & inserted.Paragraphs(1).Range.Style.NameLocal
    VTFinishPureNativeEquation inserted, id
    report = report & "number-colour-style-size=PASS" & vbLf
    stage = "bold-number-slot"
    inserted.Font.Bold = True
    VTSetEquationNumberingFormat doc, VT_WORD_NUMBERING_MODE_CHAPTER, "."
    Set inserted = VTEnsurePureNativeNumberFormat(inserted)
    VTUpdateCurrentWordNumberFields doc
    If InStr(doc.Content.Text, VT_WORD_NATIVE_NUMBER_CACHE_PLACEHOLDER) > 0 Then _
        Err.Raise vbObjectError + 7607, , "A staged number leaked into the document."
    report = report & "bold-number-slot=PASS" & vbLf
    stage = "pure-native"
    For Each mark In doc.Bookmarks
        If Left$(mark.Name, 3) = "VT_" Then _
            Err.Raise vbObjectError + 7607, , "Private native bookmark remains."
    Next mark
    For Each variable In doc.Variables
        If Left$(variable.Name, 6) = "VT_LX_" Or _
           Left$(variable.Name, 6) = "VT_MD_" Then _
            Err.Raise vbObjectError + 7607, , "Private native payload remains."
    Next variable
    doc.SaveAs2 FileName:=VTNativeParityScratchRoot() & "/native-parity.docx", _
        FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
    report = report & "pure-native=PASS" & vbLf & "status=PASS" & vbLf
Finished:
    On Error Resume Next
    VT_NATIVE_PARITY_DIAGNOSTIC = False
    VTEndWordInternalMutation
    VTWriteTextAtomic VTApplicationSupportRoot() & "/Tests/native-parity.txt", _
        report & "stage=" & stage & vbLf & "error=" & CStr(failureNumber) & _
        ":" & failureDescription
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Resume Finished
End Sub

Public Sub VisualTeX_RunUserNativeParityRegression()
    Dim doc As Document
    Dim math As OMath
    Dim exact As Range
    Dim inserted As Range
    Dim insertion As Range
    Dim body As Range
    Dim root As String
    Dim fixture As String
    Dim payload As String
    Dim originalText As String
    Dim id As String
    Dim stage As String
    Dim report As String
    Dim index As Long
    Dim bodyStart As Long
    Dim bodyEnd As Long
    Dim originalCount As Long
    Dim originalColour As Long
    Dim originalStyle As String
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    VTBeginWordInternalMutation
    root = VTNativeParityScratchRoot()
    fixture = VTApplicationSupportRoot() & "/Tests/numbering-edit/off.docx"
    payload = VTBase64UrlEncodeUtf8("<m:oMath xmlns:m=""http:" & "//schemas.openxmlformats.org/officeDocument/2006/math""><m:r><m:t>x</m:t></m:r></m:oMath>")
    stage = "user-15-open"
    Set doc = Documents.Open(FileName:=root & "/user-15-before.docx", AddToRecentFiles:=False)
    originalCount = doc.OMaths.Count
    Set math = VTNativeMathRoot(doc.OMaths(3))
    originalText = VTXmlTextContent(VTNativeMathXmlBody(math.Range.WordOpenXML, bodyStart, bodyEnd))
    For index = 1 To 3
        stage = "user-15-adjacent-" & CStr(index)
        Set insertion = doc.Range(math.Range.End, math.Range.End)
        Set inserted = VTInsertNativeEquationAtRange(insertion, payload, fixture, "inline", False, False)
        If doc.OMaths.Count <> originalCount Then Err.Raise 5, , "Native math zone count changed."
        If Left$(VTXmlTextContent(VTNativeMathXmlBody(inserted.WordOpenXML, bodyStart, bodyEnd)), Len(originalText)) <> originalText Then Err.Raise 5, , "Original mathematics changed."
        id = VTNewUuidV4()
        VTFinishPureNativeEquation inserted, id
        Set math = VTNativeMathRoot(inserted.OMaths(1))
        VTPlaceCaretAfterInlineNativeEquation inserted
    Next index
    Selection.TypeText "TAIL"
    If InStr(math.Range.Text, "TAIL") > 0 Then Err.Raise 5, , "Body text entered mathematics."
    doc.SaveAs2 FileName:=root & "/user-15-native-fixed.docx", FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
    report = report & "user-15-native-adjacency=PASS" & vbLf
    doc.Close SaveChanges:=wdDoNotSaveChanges
    stage = "user-50-open"
    Set doc = Documents.Open(FileName:=root & "/user-50-before.docx", AddToRecentFiles:=False)
    originalCount = doc.OMaths.Count
    For index = 1 To originalCount
        Set math = doc.OMaths(index)
        VTPrepareLegacyNativeMathForEdit math
        If StrComp(math.Range.Paragraphs(1).Range.Style.NameLocal, _
           doc.Styles(wdStyleCaption).NameLocal, vbTextCompare) = 0 Then _
            Err.Raise 5, , "The legacy native Caption style remains."
    Next index
    If InStr(doc.Content.Text, ChrW(8288)) > 0 Then Err.Raise 5, , "The legacy native caret boundary remains."
    Set exact = doc.OMaths(4).Range.Duplicate
    originalColour = VTNativeParityFirstGlyph(exact).Font.Color
    originalStyle = exact.Paragraphs(1).Range.Style.NameLocal
    stage = "user-50-add-number"
    id = VTNewUuidV4()
    Set inserted = VTWritePureNativeNumber(exact, id)
    Set body = VTNativeParityFirstGlyph(VTNativeEquationFormulaContentRange(inserted))
    If body.Font.Color <> originalColour Then Err.Raise 5, , "The black formula changed colour."
    If inserted.Paragraphs(1).Range.Style.NameLocal <> originalStyle Then Err.Raise 5, , "Numbering changed the source paragraph style."
    stage = "user-50-reconcile"
    VTReconcileNumberingAfterEdit doc, inserted.Start, False
    VTFinishPureNativeEquation inserted, id
    If InStr(doc.Content.Text, VT_WORD_NATIVE_NUMBER_CACHE_PLACEHOLDER) > 0 Then Err.Raise 5, , "Number sentinel remains."
    If doc.OMaths.Count <> originalCount Then Err.Raise 5, , "Original math disappeared."
    doc.SaveAs2 FileName:=root & "/user-50-native-fixed.docx", FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
    doc.Close SaveChanges:=wdDoNotSaveChanges
    Set doc = Documents.Open(FileName:=root & "/user-50-native-fixed.docx", _
        ReadOnly:=False, AddToRecentFiles:=False)
    stage = "user-50-reopen-number-state"
    If VTPureNativeNumberField(doc.OMaths(4).Range) Is Nothing Then _
        Err.Raise 5, , "The native number was not recognized after save/reopen."
    report = report & "user-50-number-colour-reconcile=PASS" & vbLf & "status=PASS" & vbLf
Finished:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=wdDoNotSaveChanges
    VTEndWordInternalMutation
    VTWriteTextAtomic VTApplicationSupportRoot() & "/Tests/user-native-parity.txt", report & "stage=" & stage & vbLf & "error=" & CStr(failureNumber) & ":" & failureDescription
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Resume Finished
End Sub
