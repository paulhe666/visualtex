Attribute VB_Name = "VTWordAdapter"
Option Explicit

Private Const VT_WORD_HOST As String = "word"
Private Const VT_WORD_STATUS_FILE As String = "/OfficePluginStatus/word.json"
Private Const VT_WORD_SOURCE_REVISION As String = _
    "word-events-external-seq-safe-insert-20260722-r31"
Private Const VT_WORD_BOOKMARK_PREFIX As String = "VT_Pending_"
Private Const VT_WORD_NATIVE_BOOKMARK_PREFIX As String = "VT_F_"
Private Const VT_WORD_CAPTION_BOOKMARK_PREFIX As String = "VT_C_"
Private Const VT_WORD_NUMBER_BOOKMARK_PREFIX As String = "VT_R_"
Private Const VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX As String = "VT_N_"
Private Const VT_WORD_LATEX_VARIABLE_PREFIX As String = "VT_Latex_"
Private Const VT_WORD_OMML_VARIABLE_PREFIX As String = "VT_OMML_"
Private Const VT_WORD_METADATA_VARIABLE_PREFIX As String = "VT_Metadata_"
Private Const VT_WORD_FORMAT_VARIABLE_PREFIX As String = "VT_Format_"
Private Const VT_WORD_PAYLOAD_CHUNK_SIZE As Long = 20000
Private Const VT_WORD_PAYLOAD_MAX_CHUNKS As Long = 128
Private Const VT_WORD_TRACE_ENABLED As Boolean = False
Private Const VT_WORD_LATEX_CHUNK_SIZE As Long = VT_WORD_PAYLOAD_CHUNK_SIZE
Private Const VT_WORD_LATEX_MAX_CHUNKS As Long = VT_WORD_PAYLOAD_MAX_CHUNKS
Private Const VT_WORD_OMML_CHUNK_SIZE As Long = VT_WORD_PAYLOAD_CHUNK_SIZE
Private Const VT_WORD_OMML_MAX_CHUNKS As Long = VT_WORD_PAYLOAD_MAX_CHUNKS
Private VT_WORD_EVENT_SINK As VTWordEvents
Private VT_WORD_INTERNAL_MUTATION_DEPTH As Long
Private VT_WORD_ORPHAN_WATCH_SCHEDULED As Boolean
Private VT_WORD_ORPHAN_WATCH_RUNNING As Boolean

Public Sub AutoExec()
    On Error Resume Next
    VTInitializeWordEvents
    VTEnsureOrphanWatchScheduled
    VTWriteWordHealth
    On Error GoTo 0
End Sub

Public Function VTWordSourceSelfTest() As Boolean
    Dim equationLabelName As String

    If VTBase64UrlDecodeUtf8("XGZyYWN7YX17Yn0") <> "\frac{a}{b}" Then Exit Function
    If VTLaTeXToWordLinear("\frac{a_1}{\sqrt{x}}") <> "(a_(1))/(\sqrt(x))" Then Exit Function
    If VTLaTeXToWordLinear("\alpha+\beta") <> "\alpha+\beta" Then Exit Function
    equationLabelName = VTNativeEquationLabelName()
    If Len(equationLabelName) = 0 Then Exit Function
    If InStr(1, VTEquationSequenceFieldText(equationLabelName), equationLabelName, vbTextCompare) = 0 Then Exit Function
    VTWordSourceSelfTest = True
End Function

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

Private Function VTAppendRegressionParagraph( _
    ByVal documentObject As Document) As Range

    Dim insertionStart As Long

    If documentObject Is Nothing Then
        Err.Raise vbObjectError + 7548, "VisualTeX", _
            "The Word regression document is missing."
    End If
    documentObject.Content.InsertAfter vbCr
    insertionStart = documentObject.Content.End - 1
    Set VTAppendRegressionParagraph = documentObject.Range( _
        Start:=insertionStart, End:=insertionStart)
End Function

Private Sub VTRegressionAssertImageNumberVerticalAlignment( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal expectedOrdinal As Long, _
    ByVal stageName As String, _
    ByRef formulaBaselineY As Single, _
    ByRef numberBaselineY As Single, _
    ByRef formulaCenterY As Single, _
    ByRef numberCenterY As Single, _
    ByRef numberPosition As Long)

    Dim formulaRange As Range
    Dim paragraphRange As Range
    Dim helperParagraph As Range
    Dim formulaShape As InlineShape
    Dim sequenceField As Field
    Dim visibleNumberField As Field
    Dim openingRange As Range
    Dim resultRange As Range
    Dim closingRange As Range
    Dim prefixRange As Range
    Dim separatorRange As Range
    Dim suffixRange As Range
    Dim formulaProbe As Range
    Dim numberProbe As Range
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim formulaLine As Long
    Dim numberLine As Long
    Dim numberLineHeight As Single
    Dim centerError As Single

    If documentObject Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Or _
       expectedOrdinal < 1 Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & ": the image-number alignment target is invalid."
    End If
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
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    Set visibleNumberField = VTImageEquationReferenceField( _
        formulaRange, formulaId)
    If sequenceField Is Nothing Or visibleNumberField Is Nothing Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & ": the external SEQ or visible REF is missing."
    End If
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If Not VTHelperParagraphOwnsNativeEquationSequence( _
           helperParagraph) Or _
       helperParagraph.Start < paragraphRange.End Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & ": the image SEQ helper paragraph is invalid."
    End If
    fieldStart = VTEquationFieldStart(visibleNumberField)
    fieldEnd = VTEquationFieldEnd(visibleNumberField)
    Set openingRange = documentObject.Range( _
        Start:=fieldStart - 1, End:=fieldStart)
    Set resultRange = visibleNumberField.Result.Duplicate
    Set closingRange = documentObject.Range( _
        Start:=fieldEnd, End:=fieldEnd + 1)
    Set prefixRange = documentObject.Range( _
        Start:=paragraphRange.Start, End:=formulaRange.Start)
    Set separatorRange = documentObject.Range( _
        Start:=formulaRange.End, End:=fieldStart)
    Set suffixRange = documentObject.Range( _
        Start:=fieldEnd, End:=paragraphRange.End)

    If paragraphRange.Paragraphs.Count <> 1 Or _
       paragraphRange.Fields.Count <> 1 Or _
       prefixRange.Text <> vbTab Or _
       separatorRange.Text <> vbTab & "(" Or _
       suffixRange.Text <> ")" & vbCr Or _
       VTEquationSequenceResultText(sequenceField) <> _
           CStr(expectedOrdinal) Or _
       VTEquationSequenceResultText(visibleNumberField) <> _
           CStr(expectedOrdinal) Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & _
            ": the layout is not <TAB><image><TAB>(REF)<CR> plus external SEQ."
    End If
    If paragraphRange.ParagraphFormat.TabStops.Count <> 2 Or _
       paragraphRange.ParagraphFormat.TabStops(1).Alignment <> _
           wdAlignTabCenter Or _
       paragraphRange.ParagraphFormat.TabStops(2).Alignment <> _
           wdAlignTabRight Or _
       Abs(paragraphRange.ParagraphFormat.TabStops(1).Position - _
           207.65!) > 0.2 Or _
       Abs(paragraphRange.ParagraphFormat.TabStops(2).Position - _
           414.3!) > 0.2 Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & _
            ": the 207.65/414.30 point tab stops changed."
    End If
    If Abs(paragraphRange.ParagraphFormat.LeftIndent) > 0.05 Or _
       Abs(paragraphRange.ParagraphFormat.RightIndent) > 0.05 Or _
       Abs(paragraphRange.ParagraphFormat.FirstLineIndent) > 0.05 Or _
       Abs(paragraphRange.ParagraphFormat.LineSpacing - 12!) > 0.1 Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & _
            ": the zero-indent, 12-point line geometry changed."
    End If
    If Abs(formulaShape.Width - 54.25!) > 0.1 Or _
       Abs(formulaShape.Height - 46.1!) > 0.1 Or _
       formulaShape.Range.Font.Position <> 0 Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & _
            ": the 54.25 x 46.10 point image fixture changed."
    End If
    If Abs(openingRange.Font.Size - 10!) > 0.1 Or _
       Abs(resultRange.Font.Size - 10!) > 0.1 Or _
       Abs(closingRange.Font.Size - 10!) > 0.1 Or _
       openingRange.Font.Position <> resultRange.Font.Position Or _
       openingRange.Font.Position <> closingRange.Font.Position Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & _
            ": (, REF result and ) do not share 10-point formatting."
    End If
    numberPosition = openingRange.Font.Position
    If numberPosition = wdUndefined Or _
       numberPosition < 16 Or numberPosition > 18 Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & _
            ": the image number Position is not approximately +17 points" & _
            " [position=" & CStr(numberPosition) & "]."
    End If

    documentObject.Repaginate
    Set formulaProbe = formulaShape.Range.Duplicate
    formulaProbe.Collapse wdCollapseStart
    Set numberProbe = openingRange.Duplicate
    numberProbe.Collapse wdCollapseStart
    formulaLine = formulaProbe.Information(wdFirstCharacterLineNumber)
    numberLine = numberProbe.Information(wdFirstCharacterLineNumber)
    If formulaLine <= 0 Or numberLine <= 0 Or _
       formulaLine <> numberLine Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & _
            ": the image and Equation number are not on one visual line."
    End If
    formulaBaselineY = CSng(formulaProbe.Information( _
        wdVerticalPositionRelativeToPage))
    numberBaselineY = CSng(numberProbe.Information( _
        wdVerticalPositionRelativeToPage))
    numberLineHeight = paragraphRange.ParagraphFormat.LineSpacing
    formulaCenterY = formulaBaselineY - formulaShape.Height / 2!
    numberCenterY = numberBaselineY - numberLineHeight / 2! - _
        numberPosition
    centerError = Abs(formulaCenterY - numberCenterY)
    If formulaBaselineY < 0! Or numberBaselineY < 0! Or _
       centerError > 2! Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            stageName & _
            ": baseline-derived visual centers differ by more than two points" & _
            " [formulaBaseline=" & CStr(formulaBaselineY) & _
            "; numberBaseline=" & CStr(numberBaselineY) & _
            "; formulaCenter=" & CStr(formulaCenterY) & _
            "; numberCenter=" & CStr(numberCenterY) & _
            "; error=" & CStr(centerError) & "]."
    End If
End Sub

Public Sub VisualTeX_RunWordImageNumberVerticalAlignmentRegression()
    Const formulaId As String = _
        "28282828-2828-4828-8828-282828282828"

    Dim sourceDocument As Document
    Dim testDocument As Document
    Dim formulaShape As InlineShape
    Dim formulaRange As Range
    Dim paragraphRange As Range
    Dim insertionRange As Range
    Dim sequenceField As Field
    Dim visibleNumberField As Field
    Dim numberRange As Range
    Dim encodedMetadata As String
    Dim latexBase64 As String
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim positionBefore As Long
    Dim positionAfter As Long
    Dim formulaBaselineBefore As Single
    Dim numberBaselineBefore As Single
    Dim formulaCenterBefore As Single
    Dim numberCenterBefore As Single
    Dim formulaBaselineAfter As Single
    Dim numberBaselineAfter As Single
    Dim formulaCenterAfter As Single
    Dim numberCenterAfter As Single

    On Error GoTo RegressionFailed
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/word-image-number-vertical-alignment-regression-result.txt"
    If Documents.Count > 0 Then Set sourceDocument = ActiveDocument

    regressionStage = "create-54p25-by-46p10-fixture"
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.Activate
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
    Set visibleNumberField = VTImageEquationReferenceField( _
        formulaRange, formulaId)
    If sequenceField Is Nothing Or visibleNumberField Is Nothing Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The image-number regression SEQ/REF identity is missing."
    End If
    fieldStart = VTEquationFieldStart(visibleNumberField)
    fieldEnd = VTEquationFieldEnd(visibleNumberField)
    Set numberRange = testDocument.Range( _
        Start:=fieldStart - 1, End:=fieldEnd + 1)
    With numberRange.Font
        .Size = 10!
        .Position = 0
    End With
    formulaRange.InlineShapes(1).Range.Font.Position = 0
    VTSetEquationNumberBookmarkExact _
        testDocument, formulaId, numberRange
    positionBefore = VTCalibrateImageEquationNumberPosition( _
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
        "before field refresh", _
        formulaBaselineBefore, numberBaselineBefore, _
        formulaCenterBefore, numberCenterBefore, positionBefore

    regressionStage = "refresh-fields-and-repaginate"
    testDocument.Fields.Update
    VTReconcileEquationNumbers testDocument
    testDocument.Repaginate
    Set formulaRange = VTNumberedFormulaRangeForId( _
        testDocument, formulaId)
    VTAssertNumberedEquationLayout _
        formulaRange, 46.1!, formulaId, _
        "vertical alignment fixture", _
        "image-number baseline regression after field refresh"
    VTVerifyParagraphEquationNumberIntegrity _
        formulaRange, formulaId, 1
    VTRegressionAssertImageNumberVerticalAlignment _
        testDocument, formulaId, 1, _
        "after field refresh", _
        formulaBaselineAfter, numberBaselineAfter, _
        formulaCenterAfter, numberCenterAfter, positionAfter

    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    VTWriteTextAtomic resultPath, _
        "PASS" & vbLf & _
        "revision=" & VT_WORD_SOURCE_REVISION & vbLf & _
        "structure=<TAB><image><TAB>(REF)<CR>+external-SEQ" & vbLf & _
        "image=54.25x46.10" & vbLf & _
        "fontSize=10" & vbLf & _
        "lineSpacing=12" & vbLf & _
        "centerTab=207.65" & vbLf & _
        "rightTab=414.30" & vbLf & _
        "positionBefore=" & CStr(positionBefore) & vbLf & _
        "positionAfter=" & CStr(positionAfter) & vbLf & _
        "formulaBaselineBefore=" & _
            CStr(formulaBaselineBefore) & vbLf & _
        "numberBaselineBefore=" & CStr(numberBaselineBefore) & vbLf & _
        "formulaCenterBefore=" & CStr(formulaCenterBefore) & vbLf & _
        "numberCenterBefore=" & CStr(numberCenterBefore) & vbLf & _
        "centerErrorBefore=" & _
            CStr(Abs(formulaCenterBefore - numberCenterBefore)) & vbLf & _
        "formulaBaselineAfter=" & _
            CStr(formulaBaselineAfter) & vbLf & _
        "numberBaselineAfter=" & CStr(numberBaselineAfter) & vbLf & _
        "formulaCenterAfter=" & CStr(formulaCenterAfter) & vbLf & _
        "numberCenterAfter=" & CStr(numberCenterAfter) & vbLf & _
        "centerErrorAfter=" & _
            CStr(Abs(formulaCenterAfter - numberCenterAfter)) & vbLf
    Exit Sub

RegressionFailed:
    regressionErrorNumber = Err.Number
    regressionErrorDescription = Err.Description
    On Error Resume Next
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    VTWriteTextAtomic resultPath, _
        "FAIL" & vbLf & _
        "revision=" & VT_WORD_SOURCE_REVISION & vbLf & _
        "stage=" & regressionStage & vbLf & _
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & _
            Replace$(Replace$(regressionErrorDescription, vbCr, " "), _
                vbLf, " ") & vbLf
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
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
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
    Dim fixtureRoot As String
    Dim nativeDocumentPath As String
    Dim ommlBase64 As String
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
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
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
    ByVal heightPoints As Single) As InlineShape

    Dim insertionRange As Range
    Dim numberRange As Range
    Dim formulaShape As InlineShape
    Dim encodedMetadata As String

    encodedMetadata = VT_METADATA_PREFIX & "e30"
    Set insertionRange = VTAppendRegressionParagraph(documentObject)
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
    Set numberRange = VTInsertEquationNumber( _
        formulaShape, formulaId, _
        VTEquationCrossReferenceText(latexBase64))
    Set numberRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
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

Private Function VTRegressionCreateNumberedNative( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal latexBase64 As String, _
    ByVal ommlBase64 As String, _
    ByVal nativeDocumentPath As String) As Range

    Dim insertionRange As Range
    Dim equationRange As Range
    Dim numberRange As Range
    Dim encodedMetadata As String
    Dim numberCreated As Boolean

    encodedMetadata = VT_METADATA_PREFIX & "e30"
    Set insertionRange = VTAppendRegressionParagraph(documentObject)
    Set equationRange = VTInsertNativeEquationAtRange( _
        insertionRange, ommlBase64, nativeDocumentPath, _
        "inline", True, False)
    VTSetWordLatexPayload documentObject, formulaId, latexBase64
    VTSetWordOmmlPayload documentObject, formulaId, ommlBase64
    VTSetWordMetadataPayload documentObject, formulaId, encodedMetadata
    VTSetWordFormulaFormat documentObject, formulaId, "block", True
    Set numberRange = VTEnsureNativeEquationNumber( _
        equationRange, 48#, formulaId, _
        VTEquationCrossReferenceText(latexBase64), numberCreated)
    If Not numberCreated Then
        Err.Raise vbObjectError + 7554, "VisualTeX", _
            "The deletion regression native formula did not create a number."
    End If
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
    Dim nativeStart As Long
    Dim numberCreated As Boolean

    If documentObject Is Nothing Or requestedRange Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The safe native insertion regression target is invalid."
    End If
    encodedMetadata = VT_METADATA_PREFIX & "e30"
    Set insertionRange = VTPrepareWordCreateInsertionRange( _
        requestedRange.Duplicate, "block")
    Set placeholder = documentObject.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, SaveWithDocument:=True, _
        Range:=insertionRange)
    placeholder.Width = 1!
    placeholder.Height = 1!

    Set equationRange = VTInsertNativeEquationAtRange( _
        placeholder.Range.Duplicate, ommlBase64, nativeDocumentPath, _
        "inline", True, False)
    nativeStart = equationRange.Start
    placeholder.Delete
    Set equationRange = VTResolveNativeEquationRange( _
        documentObject, nativeStart, 16)
    VTSetNativeFormulaBookmark documentObject, equationRange, formulaId

    VTSetWordLatexPayload documentObject, formulaId, latexBase64
    VTSetWordOmmlPayload documentObject, formulaId, ommlBase64
    VTSetWordMetadataPayload documentObject, formulaId, encodedMetadata
    VTSetWordFormulaFormat documentObject, formulaId, "block", True
    Set numberRange = VTEnsureNativeEquationNumber( _
        equationRange, 48#, formulaId, _
        VTEquationCrossReferenceText(latexBase64), numberCreated)
    If Not numberCreated Then
        Err.Raise vbObjectError + 7566, "VisualTeX", _
            "The safe native insertion did not create its own Equation number."
    End If

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
    VTSetNativeFormulaBookmark documentObject, equationRange, formulaId
    Set VTRegressionCreateNumberedNativeAtCaret = equationRange.Duplicate
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
    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Content.Text = "SAFE_BODY_ABOVE"
    Set existingRange = VTRegressionCreateNumberedNative( _
        testDocument, bodyGapExistingId, "eF8y", ommlBase64, _
        nativeDocumentPath)
    Set blankRange = VTRegressionInsertBlankParagraphBeforeFormula( _
        testDocument, bodyGapExistingId)
    Set formulaRange = VTRegressionCreateNumberedNativeAtCaret( _
        testDocument, blankRange, bodyGapFormulaId, "eF8z", _
        ommlBase64, nativeDocumentPath)
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
        "errorNumber=" & CStr(regressionErrorNumber) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            regressionErrorDescription, vbCr, " "), vbLf, " ") & vbLf
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word safe native insertion regression", _
        regressionStage & ": " & regressionErrorDescription
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
               Not candidateField.Result.Information(wdWithInTable) Then
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
           Not candidateField.Result.Information(wdWithInTable) Then
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
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX Word deletion/reference regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Private Sub VTLegacyWordReferencePersistenceRegression()
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
    Dim builtInBookmarkRange As Range
    Dim candidateField As Field
    Dim ommlBase64 As String
    Dim nativeDocumentPath As String
    Dim fixtureRoot As String
    Dim resultPath As String
    Dim customTargetName As String
    Dim builtInTargetName As String
    Dim fieldTargetName As String
    Dim customResult As String
    Dim builtInResult As String
    Dim bodyFieldCountBefore As Long
    Dim bodyFieldCountAfter As Long
    Dim newestBodyFieldStart As Long
    Dim candidateFieldStart As Long
    Dim capturedBindings As Collection
    Dim bindingItem As Variant
    Dim builtInBindingCaptured As Boolean
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    fixtureRoot = VTApplicationSupportRoot() & "/Tests"
    nativeDocumentPath = _
        VTApplicationSupportRoot() & "/NativeDocuments/" & _
        "11111111-1111-4111-8111-111111111111.docx"
    resultPath = fixtureRoot & _
        "/word-reference-persistence-regression-result.txt"
    ommlBase64 = VTReadText( _
        fixtureRoot & "/word-native-regression-omml.txt", _
        VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS)
    If Not VTPathFileExists(nativeDocumentPath) Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "The reference persistence native DOCX fixture is missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate

    regressionStage = "create-two-numbered-images"
    Set firstFormula = VTRegressionCreateNumberedImage( _
        testDocument, firstFormulaId, "eF8x", 120!, 34!)
    Set referencedFormula = VTRegressionCreateNumberedImage( _
        testDocument, referencedFormulaId, "eF8y", 130!, 38!)
    VTSetWordOmmlPayload testDocument, referencedFormulaId, ommlBase64
    customTargetName = _
        VTEquationSequenceNumberBookmarkName(referencedFormulaId)

    regressionStage = "insert-visualtex-reference"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    Set insertedReference = VTInsertEquationNumberReferenceAtRange( _
        insertionRange, 2)
    If insertedReference.Text <> "(2)" Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "The VisualTeX reference was not inserted as (2)."
    End If

    regressionStage = "insert-word-built-in-reference"
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
            candidateFieldStart = VTEquationFieldStart(candidateField)
            If candidateFieldStart > newestBodyFieldStart Then
                newestBodyFieldStart = candidateFieldStart
                fieldTargetName = VTReferenceTargetBookmarkName( _
                    candidateField.Code.Text)
                builtInTargetName = fieldTargetName
            End If
        End If
    Next candidateField
    If Len(builtInTargetName) = 0 Or _
       bodyFieldCountAfter <> bodyFieldCountBefore + 1 Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "Word did not create a built-in Equation REF target."
    End If

    regressionStage = "capture-word-built-in-reference-binding"
    Set capturedBindings = _
        VTCaptureBodyEquationReferenceBindings(testDocument)
    For Each bindingItem In capturedBindings
        If InStr(1, CStr(bindingItem), _
           builtInTargetName & vbTab & referencedFormulaId & vbTab, _
           vbTextCompare) = 1 Then
            builtInBindingCaptured = True
            Exit For
        End If
    Next bindingItem
    If Not builtInBindingCaptured Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "Word built-in REF was not bound to its live formula before" & _
            " renumbering [target=" & builtInTargetName & "]."
    End If

    regressionStage = "delete-preceding-formula-and-renumber"
    Set insertionRange = VTNumberedFormulaRangeForId( _
        testDocument, firstFormulaId)
    insertionRange.InlineShapes(1).Delete
    VTPruneOrphanedEquationNumberScaffolds testDocument
    VTReconcileEquationNumbers testDocument
    customResult = VTBodyReferenceResultForTarget( _
        testDocument, customTargetName)
    builtInResult = VTBodyReferenceResultForTarget( _
        testDocument, builtInTargetName)
    If customResult <> "1" Or InStr(1, builtInResult, "1", _
       vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "A surviving Equation reference disappeared after renumbering" & _
            " [visualtex=" & customResult & _
            "; builtIn=" & builtInResult & "]."
    End If

    regressionStage = "convert-referenced-image-to-native"
    Set insertionRange = VTNumberedFormulaRangeForId( _
        testDocument, referencedFormulaId)
    Set referencedFormula = insertionRange.InlineShapes(1)
    VTWordConvertInlineShapeToNativeEquation referencedFormula
    customResult = VTBodyReferenceResultForTarget( _
        testDocument, customTargetName)
    builtInResult = VTBodyReferenceResultForTarget( _
        testDocument, builtInTargetName)
    If customResult <> "1" Or InStr(1, builtInResult, "1", _
       vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "A live Equation reference disappeared during image-to-OMML" & _
            " conversion [visualtex=" & customResult & _
            "; builtIn=" & builtInResult & "]."
    End If
    If Not VTTryGetBookmarkRangeIncludingHidden( _
       testDocument, builtInTargetName, builtInBookmarkRange) Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "The Word built-in reference target Bookmark was not restored."
    End If

    testDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set testDocument = Nothing
    VTWriteTextAtomic resultPath, _
        "PASS" & vbLf & _
        "visualTeXReference=1" & vbLf & _
        "wordBuiltInReference=" & builtInResult & vbLf & _
        "imageToOmmlReference=1" & vbLf
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
        "VisualTeX Word reference persistence regression", _
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
           Not candidateField.Result.Information(wdWithInTable) Then
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
           Not candidateField.Result.Information(wdWithInTable) Then
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
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
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
    Dim crossReferenceInventory As String
    Dim equationEndBefore As Long
    Dim equationEndAfter As Long
    Dim caretPositionBefore As Long
    Dim diagnosticIndex As Long
    Dim characterCode As Long
    Dim sequenceField As Field
    Dim referenceField As Field
    Dim newParagraph As Paragraph
    Dim diagnosticRange As Range
    Dim crossReferenceTextFound As Boolean
    Dim crossReferenceNumberFound As Boolean
    Dim referenceResult As String
    Dim invariantSnapshot As String
    Dim invariantParagraphStart As Long
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
       Not VTPathFileExists(integralDocumentPath) Or _
       Not VTPathFileExists(sumFractionDocumentPath) Then
        Err.Raise vbObjectError + 7482, "VisualTeX", _
            "A Word native regression DOCX fixture is missing."
    End If

    Set testDocument = Documents.Add(Visible:=True)
    testDocument.ActiveWindow.View.Type = wdPrintView
    testDocument.Activate

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
    Set insertionRange = testDocument.Range( _
        Start:=testDocument.Content.End - 1, _
        End:=testDocument.Content.End - 1)
    insertionRange.InsertBefore vbCr
    Set insertionRange = testDocument.Range( _
        Start:=testDocument.Content.End - 1, _
        End:=testDocument.Content.End - 1)
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
    referenceResult = Trim$(diagnosticRange.Text)
    If referenceResult <> "1" Or _
       Trim$(referenceField.Result.Text) <> "1" Or _
       diagnosticRange.InlineShapes.Count <> 0 Or _
       InStr(1, referenceResult, "(", vbBinaryCompare) > 0 Or _
       InStr(1, referenceResult, ")", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "Word's built-in image Equation reference is not a pure number" & _
            " [result=" & referenceField.Result.Text & _
            "; text=" & referenceResult & _
            "; images=" & CStr(diagnosticRange.InlineShapes.Count) & "]."
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
    Set numberRange = VTInsertEquationNumber( _
        referenceStabilityPlaceholder, stabilityImageFormulaId, _
        "later reference stability formula")
    VTReconcileEquationNumbers testDocument

    Set referenceField = Nothing
    For Each sequenceField In testDocument.Fields
        If sequenceField.Type = wdFieldRef And _
           Not sequenceField.Result.Information(wdWithInTable) Then
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
    If diagnosticRange.ParagraphFormat.Alignment <> wdAlignParagraphCenter Or _
       VTCustomTabStopCount(diagnosticRange) <> 0 Then
        Err.Raise vbObjectError + 7524, "VisualTeX", _
            "The unnumbered image display formula is not centered" & _
            " [alignment=" & _
            CStr(diagnosticRange.ParagraphFormat.Alignment) & _
            "; tabs=" & _
            CStr(diagnosticRange.ParagraphFormat.TabStops.Count) & _
            "; customTabs=" & CStr(VTCustomTabStopCount(diagnosticRange)) & _
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
    testDocument.Content.Delete
    Set insertionRange = testDocument.Range(Start:=0, End:=0)
    Set placeholder = testDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Range:=insertionRange)
    placeholder.Width = 180
    placeholder.Height = 64
    sourceHeightPoints = placeholder.Height
    Set numberRange = VTInsertEquationNumber( _
        placeholder, conversionFormulaId, "image conversion formula")
    VTAssertNumberedEquationLayout _
        placeholder.Range.Duplicate, sourceHeightPoints, conversionFormulaId, _
        "image conversion formula", "conversion-image-before"
    Set sequenceField = VTFindEquationSequenceField( _
        placeholder.Range.Paragraphs(1).Range)
    previousNumberText = sequenceField.Result.Text
    Set insertionRange = placeholder.Range.Duplicate
    insertionRange.Collapse wdCollapseEnd
    Set equationRange = VTInsertNativeEquationAtRange( _
        insertionRange, ommlBase64, nativeDocumentPath, _
        "inline", True, False)
    nativeEquationStart = equationRange.Start
    placeholder.Delete
    Set equationRange = VTResolveNativeEquationRange( _
        testDocument, nativeEquationStart, 16)
    VTSetWordOmmlPayload testDocument, conversionFormulaId, ommlBase64
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
       sequenceField.Result.Text <> previousNumberText Then
        Err.Raise vbObjectError + 7529, "VisualTeX", _
            "Image-to-native conversion changed its Equation number."
    End If

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
    Set newParagraph = testDocument.Paragraphs.Add
    Set insertionRange = newParagraph.Range.Duplicate
    insertionRange.Collapse wdCollapseStart
    Set placeholder = testDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), LinkToFile:=False, _
        SaveWithDocument:=True, Range:=insertionRange)
    placeholder.Width = 130
    placeholder.Height = 36
    Set numberRange = VTInsertEquationNumber( _
        placeholder, stabilityImageFormulaId, "later image formula")
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
    If Not testDocument Is Nothing Then
        testDocument.Close SaveChanges:=wdDoNotSaveChanges
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
            measurementDocument.OMaths(1).Range.Duplicate)
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

Private Sub VTBeginWordInternalMutation()
    VT_WORD_INTERNAL_MUTATION_DEPTH = _
        VT_WORD_INTERNAL_MUTATION_DEPTH + 1
End Sub

Private Sub VTEndWordInternalMutation()
    If VT_WORD_INTERNAL_MUTATION_DEPTH > 0 Then
        VT_WORD_INTERNAL_MUTATION_DEPTH = _
            VT_WORD_INTERNAL_MUTATION_DEPTH - 1
    End If
End Sub

Public Function VTWordInternalMutationActive() As Boolean
    VTWordInternalMutationActive = _
        (VT_WORD_INTERNAL_MUTATION_DEPTH > 0)
End Function

Public Sub VTInitializeWordEvents()
    Set VT_WORD_EVENT_SINK = New VTWordEvents
    Set VT_WORD_EVENT_SINK.App = Word.Application
End Sub

Private Function VTNumberedDisplayTableNearRange( _
    ByVal selectedRange As Range) As Table

    Dim documentObject As Document
    Dim probeRange As Range
    Dim candidateTable As Table
    Dim probeStart As Long

    If selectedRange Is Nothing Then Exit Function
    Set documentObject = selectedRange.Document
    probeStart = selectedRange.Start

    On Error Resume Next
    If selectedRange.Information(wdWithInTable) Then
        Set candidateTable = selectedRange.Tables(1)
    End If
    If candidateTable Is Nothing And probeStart > 0 Then
        Set probeRange = documentObject.Range( _
            Start:=probeStart - 1, End:=probeStart)
        If probeRange.Information(wdWithInTable) Then
            Set candidateTable = probeRange.Tables(1)
        End If
    End If
    If candidateTable Is Nothing And _
       probeStart < documentObject.Content.End - 1 Then
        Set probeRange = documentObject.Range( _
            Start:=probeStart, End:=probeStart + 1)
        If probeRange.Information(wdWithInTable) Then
            Set candidateTable = probeRange.Tables(1)
        End If
    End If
    On Error GoTo 0

    If candidateTable Is Nothing Then Exit Function
    If candidateTable.Rows.Count <> 1 Or _
       candidateTable.Columns.Count <> 3 Then Exit Function
    Set VTNumberedDisplayTableNearRange = candidateTable
End Function

Private Function VTNumberBookmarkNameForTable( _
    ByVal layoutTable As Table) As String

    Dim documentObject As Document
    Dim candidateBookmark As Bookmark
    Dim bookmarkTable As Table

    If layoutTable Is Nothing Then Exit Function
    Set documentObject = layoutTable.Range.Document
    For Each candidateBookmark In documentObject.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_NUMBER_BOOKMARK_PREFIX Then
            Set bookmarkTable = Nothing
            On Error Resume Next
            If candidateBookmark.Range.Information(wdWithInTable) Then
                Set bookmarkTable = candidateBookmark.Range.Tables(1)
            End If
            On Error GoTo 0
            If Not bookmarkTable Is Nothing Then
                If bookmarkTable.Range.Start = layoutTable.Range.Start And _
                   bookmarkTable.Range.End = layoutTable.Range.End Then
                    VTNumberBookmarkNameForTable = candidateBookmark.Name
                    Exit Function
                End If
            End If
        End If
    Next candidateBookmark
End Function

Private Function VTFormulaIdFromBookmarkSuffix( _
    ByVal suffixText As String) As String

    Dim characterIndex As Long
    Dim characterValue As String
    Dim formulaId As String

    If Len(suffixText) <> 32 Then Exit Function
    For characterIndex = 1 To Len(suffixText)
        characterValue = Mid$(suffixText, characterIndex, 1)
        If InStr(1, "0123456789abcdefABCDEF", _
           characterValue, vbBinaryCompare) = 0 Then Exit Function
    Next characterIndex
    formulaId = _
        Left$(suffixText, 8) & "-" & _
        Mid$(suffixText, 9, 4) & "-" & _
        Mid$(suffixText, 13, 4) & "-" & _
        Mid$(suffixText, 17, 4) & "-" & _
        Right$(suffixText, 12)
    If VTIsCanonicalUuid(formulaId) Then
        VTFormulaIdFromBookmarkSuffix = formulaId
    End If
End Function

Private Sub VTCollectionAddUniqueText( _
    ByVal values As Collection, _
    ByVal value As String)

    If values Is Nothing Or Len(value) = 0 Then Exit Sub
    On Error Resume Next
    values.Add value, LCase$(value)
    On Error GoTo 0
End Sub

Private Function VTCollectionContainsText( _
    ByVal values As Collection, _
    ByVal value As String) As Boolean

    Dim candidate As Variant

    If values Is Nothing Or Len(value) = 0 Then Exit Function
    On Error Resume Next
    candidate = values(LCase$(value))
    VTCollectionContainsText = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function VTReferenceTargetBookmarkName( _
    ByVal fieldCode As String) As String

    Dim normalizedCode As String
    Dim characterIndex As Long
    Dim currentCharacter As String

    normalizedCode = Trim$(fieldCode)
    If UCase$(Left$(normalizedCode, 3)) <> "REF" Then Exit Function
    normalizedCode = LTrim$(Mid$(normalizedCode, 4))
    For characterIndex = 1 To Len(normalizedCode)
        currentCharacter = Mid$(normalizedCode, characterIndex, 1)
        If currentCharacter = " " Or currentCharacter = vbTab Then Exit For
        VTReferenceTargetBookmarkName = _
            VTReferenceTargetBookmarkName & currentCharacter
    Next characterIndex
End Function

Private Function VTFormulaIdFromSequenceBookmarkName( _
    ByVal bookmarkName As String) As String

    If Left$(bookmarkName, _
       Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)) <> _
       VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX Then Exit Function
    VTFormulaIdFromSequenceBookmarkName = VTFormulaIdFromBookmarkSuffix( _
        Mid$(bookmarkName, _
            Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX) + 1))
End Function

Private Function VTTryGetBookmarkRangeIncludingHidden( _
    ByVal documentObject As Document, _
    ByVal bookmarkName As String, _
    ByRef bookmarkRange As Range) As Boolean

    Dim previousShowHidden As Boolean

    Set bookmarkRange = Nothing
    If documentObject Is Nothing Or Len(bookmarkName) = 0 Then Exit Function
    On Error GoTo LookupFailed
    previousShowHidden = documentObject.Bookmarks.ShowHidden
    documentObject.Bookmarks.ShowHidden = True
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        Set bookmarkRange = documentObject.Bookmarks( _
            bookmarkName).Range.Duplicate
        VTTryGetBookmarkRangeIncludingHidden = True
    End If
LookupFinished:
    On Error Resume Next
    documentObject.Bookmarks.ShowHidden = previousShowHidden
    On Error GoTo 0
    Exit Function
LookupFailed:
    VTTryGetBookmarkRangeIncludingHidden = False
    Resume LookupFinished
End Function

Private Sub VTReplaceBookmarkRangeIncludingHidden( _
    ByVal documentObject As Document, _
    ByVal bookmarkName As String, _
    ByVal sourceRange As Range)

    Dim previousShowHidden As Boolean
    Dim operationErrorNumber As Long
    Dim operationErrorDescription As String

    If documentObject Is Nothing Or Len(bookmarkName) = 0 Or _
       sourceRange Is Nothing Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "The hidden Bookmark replacement target is incomplete."
    End If
    On Error GoTo ReplacementFailed
    previousShowHidden = documentObject.Bookmarks.ShowHidden
    documentObject.Bookmarks.ShowHidden = True
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        documentObject.Bookmarks(bookmarkName).Delete
    End If
    documentObject.Bookmarks.Add _
        name:=bookmarkName, Range:=sourceRange.Duplicate
    If Not documentObject.Bookmarks.Exists(bookmarkName) Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "Word did not restore its hidden cross-reference Bookmark."
    End If
ReplacementFinished:
    On Error Resume Next
    documentObject.Bookmarks.ShowHidden = previousShowHidden
    On Error GoTo 0
    If operationErrorNumber <> 0 Then
        Err.Raise operationErrorNumber, "VisualTeX", _
            operationErrorDescription
    End If
    Exit Sub
ReplacementFailed:
    operationErrorNumber = Err.Number
    operationErrorDescription = Err.Description
    Resume ReplacementFinished
End Sub

Private Function VTFirstPositiveIntegerInText( _
    ByVal sourceText As String) As Long

    Dim characterIndex As Long
    Dim currentCharacter As String
    Dim digitText As String

    For characterIndex = 1 To Len(sourceText)
        currentCharacter = Mid$(sourceText, characterIndex, 1)
        If currentCharacter >= "0" And currentCharacter <= "9" Then
            digitText = digitText & currentCharacter
        ElseIf Len(digitText) > 0 Then
            Exit For
        End If
    Next characterIndex
    If Len(digitText) > 0 Then
        VTFirstPositiveIntegerInText = CLng(Val(digitText))
    End If
End Function

Private Function VTFormulaIdForReferenceTarget( _
    ByVal documentObject As Document, _
    ByVal targetBookmarkName As String, _
    ByRef targetKind As String, _
    Optional ByVal referenceResultText As String = "", _
    Optional ByVal knownFormulaIds As Variant) As String

    Dim formulaIds As Variant
    Dim formulaId As String
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim captionBookmarkName As String
    Dim targetRange As Range
    Dim candidateRange As Range
    Dim itemIndex As Long
    Dim referenceOrdinal As Long

    targetKind = ""
    If documentObject Is Nothing Or Len(targetBookmarkName) = 0 Then
        Exit Function
    End If
    formulaId = VTFormulaIdFromSequenceBookmarkName(targetBookmarkName)
    If Len(formulaId) > 0 Then
        targetKind = "N"
        VTFormulaIdForReferenceTarget = formulaId
        Exit Function
    End If
    If IsMissing(knownFormulaIds) Then
        formulaIds = VTValidNumberedFormulaIds(documentObject)
    Else
        formulaIds = knownFormulaIds
    End If
    If Not IsArray(formulaIds) Then Exit Function
    If VTTryGetBookmarkRangeIncludingHidden( _
       documentObject, targetBookmarkName, targetRange) Then
        For itemIndex = LBound(formulaIds) To UBound(formulaIds)
            formulaId = CStr(formulaIds(itemIndex))
            sequenceBookmarkName = _
                VTEquationSequenceNumberBookmarkName(formulaId)
            numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
            captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
            If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
                Set candidateRange = documentObject.Bookmarks( _
                    sequenceBookmarkName).Range.Duplicate
                If targetRange.Start <= candidateRange.End And _
                   targetRange.End >= candidateRange.Start Then
                    targetKind = "N"
                    VTFormulaIdForReferenceTarget = formulaId
                    Exit Function
                End If
            End If
            If documentObject.Bookmarks.Exists(numberBookmarkName) Then
                Set candidateRange = documentObject.Bookmarks( _
                    numberBookmarkName).Range.Duplicate
                If targetRange.Start <= candidateRange.End And _
                   targetRange.End >= candidateRange.Start Then
                    targetKind = "R"
                    VTFormulaIdForReferenceTarget = formulaId
                    Exit Function
                End If
            End If
            If documentObject.Bookmarks.Exists(captionBookmarkName) Then
                Set candidateRange = documentObject.Bookmarks( _
                    captionBookmarkName).Range.Duplicate
                If targetRange.Start <= candidateRange.End And _
                   targetRange.End >= candidateRange.Start Then
                    targetKind = "C"
                    VTFormulaIdForReferenceTarget = formulaId
                    Exit Function
                End If
            End If
        Next itemIndex
    End If

    ' Word for Mac can create a private _Ref Bookmark whose Range is not
    ' stably nested inside the native SEQ/caption Range. Before numbering is
    ' mutated, the visible REF result still exposes its ordinal; use it as a
    ' deterministic fallback to lock the reference to the matching formula id.
    referenceOrdinal = VTFirstPositiveIntegerInText(referenceResultText)
    If referenceOrdinal >= 1 And _
       referenceOrdinal <= VTVariantArrayCount(formulaIds) Then
        targetKind = "N"
        VTFormulaIdForReferenceTarget = _
            CStr(formulaIds(LBound(formulaIds) + referenceOrdinal - 1))
    End If
End Function

Private Function VTCaptureBodyEquationReferenceBindings( _
    ByVal documentObject As Document) As Collection

    Dim bindings As New Collection
    Dim candidateField As Field
    Dim targetBookmarkName As String
    Dim formulaId As String
    Dim targetKind As String

    If documentObject Is Nothing Then
        Set VTCaptureBodyEquationReferenceBindings = bindings
        Exit Function
    End If
    For Each candidateField In documentObject.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            targetBookmarkName = VTReferenceTargetBookmarkName( _
                candidateField.Code.Text)
            formulaId = VTFormulaIdForReferenceTarget( _
                documentObject, targetBookmarkName, targetKind, _
                candidateField.Result.Text)
            If Len(formulaId) > 0 And Len(targetKind) > 0 Then
                VTCollectionAddUniqueText bindings, _
                    targetBookmarkName & vbTab & formulaId & _
                    vbTab & targetKind
            End If
        End If
    Next candidateField
    Set VTCaptureBodyEquationReferenceBindings = bindings
End Function

Private Sub VTReplaceBodyReferenceTargetWithDeletedMarker( _
    ByVal documentObject As Document, _
    ByVal targetBookmarkName As String)

    Dim candidateField As Field
    Dim replacementRange As Range
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim fieldIndex As Long

    If documentObject Is Nothing Or Len(targetBookmarkName) = 0 Then Exit Sub
    For fieldIndex = documentObject.Fields.Count To 1 Step -1
        Set candidateField = documentObject.Fields(fieldIndex)
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) And _
           StrComp(VTReferenceTargetBookmarkName( _
               candidateField.Code.Text), targetBookmarkName, _
               vbTextCompare) = 0 Then
            fieldStart = VTEquationFieldStart(candidateField)
            fieldEnd = VTEquationFieldEnd(candidateField)
            If fieldStart > 0 Then
                If documentObject.Range( _
                   Start:=fieldStart - 1, End:=fieldStart).Text = "(" Then
                    fieldStart = fieldStart - 1
                End If
            End If
            If fieldEnd < documentObject.Content.End - 1 Then
                If documentObject.Range( _
                   Start:=fieldEnd, End:=fieldEnd + 1).Text = ")" Then
                    fieldEnd = fieldEnd + 1
                End If
            End If
            Set replacementRange = documentObject.Range( _
                Start:=fieldStart, End:=fieldEnd)
            replacementRange.Text = "(deleted equation)"
            With replacementRange.Font
                .Hidden = False
                .Color = wdColorAutomatic
                .Position = 0
                .Size = VTVisibleEquationNumberFontSize(documentObject)
            End With
        End If
    Next fieldIndex
End Sub

Private Sub VTRepairBodyReferenceFieldsForTarget( _
    ByVal documentObject As Document, _
    ByVal targetBookmarkName As String, _
    ByVal expectedNumber As String)

    Dim candidateField As Field
    Dim repairedField As Field
    Dim fieldRange As Range
    Dim visibleRange As Range
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim hasOpenParenthesis As Boolean
    Dim hasCloseParenthesis As Boolean
    Dim fieldIndex As Long

    If documentObject Is Nothing Or Len(targetBookmarkName) = 0 Or _
       Len(expectedNumber) = 0 Then Exit Sub
    For fieldIndex = documentObject.Fields.Count To 1 Step -1
        Set candidateField = documentObject.Fields(fieldIndex)
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) And _
           StrComp(VTReferenceTargetBookmarkName( _
               candidateField.Code.Text), targetBookmarkName, _
               vbTextCompare) = 0 Then
            On Error Resume Next
            candidateField.Update
            On Error GoTo 0
            fieldStart = VTEquationFieldStart(candidateField)
            fieldEnd = VTEquationFieldEnd(candidateField)
            hasOpenParenthesis = False
            hasCloseParenthesis = False
            If fieldStart > 0 Then
                hasOpenParenthesis = (documentObject.Range( _
                    Start:=fieldStart - 1, End:=fieldStart).Text = "(")
            End If
            If fieldEnd < documentObject.Content.End - 1 Then
                hasCloseParenthesis = (documentObject.Range( _
                    Start:=fieldEnd, End:=fieldEnd + 1).Text = ")")
            End If
            If Trim$(candidateField.Result.Text) <> expectedNumber Then
                Set fieldRange = documentObject.Range( _
                    Start:=fieldStart, End:=fieldEnd)
                fieldRange.Text = ""
                Set fieldRange = documentObject.Range( _
                    Start:=fieldStart, End:=fieldStart)
                Set repairedField = documentObject.Fields.Add( _
                    Range:=fieldRange, Type:=wdFieldRef, _
                    Text:=VTParenthesizedEquationReferenceFieldText( _
                        targetBookmarkName), _
                    PreserveFormatting:=False)
                repairedField.Update
            Else
                Set repairedField = candidateField
            End If
            Set visibleRange = documentObject.Range( _
                Start:=IIf(hasOpenParenthesis, fieldStart - 1, fieldStart), _
                End:=VTEquationFieldEnd(repairedField) + _
                    IIf(hasCloseParenthesis, 1, 0))
            VTFormatBodyEquationReference _
                documentObject, repairedField, visibleRange
            If Trim$(repairedField.Result.Text) <> expectedNumber Then
                Err.Raise vbObjectError + 7555, "VisualTeX", _
                    "Word did not restore a live Equation body reference" & _
                    " [target=" & targetBookmarkName & _
                    "; result=" & repairedField.Result.Text & _
                    "; expected=" & expectedNumber & "]."
            End If
        End If
    Next fieldIndex
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

Private Sub VTRestoreBodyEquationReferenceBindings( _
    ByVal documentObject As Document, _
    ByVal bindings As Collection)

    Dim bindingText As String
    Dim targetBookmarkName As String
    Dim formulaId As String
    Dim targetKind As String
    Dim sourceBookmarkName As String
    Dim expectedResult As String
    Dim firstSeparator As Long
    Dim secondSeparator As Long
    Dim itemIndex As Long

    If documentObject Is Nothing Or bindings Is Nothing Then Exit Sub
    For itemIndex = 1 To bindings.Count
        bindingText = CStr(bindings(itemIndex))
        firstSeparator = InStr(1, bindingText, vbTab, vbBinaryCompare)
        secondSeparator = InStr(firstSeparator + 1, bindingText, _
            vbTab, vbBinaryCompare)
        If firstSeparator > 1 And secondSeparator > firstSeparator + 1 Then
            targetBookmarkName = Left$(bindingText, firstSeparator - 1)
            formulaId = Mid$(bindingText, firstSeparator + 1, _
                secondSeparator - firstSeparator - 1)
            targetKind = Mid$(bindingText, secondSeparator + 1)
            Select Case targetKind
                Case "R"
                    sourceBookmarkName = _
                        VTEquationNumberBookmarkName(formulaId)
                Case "C"
                    sourceBookmarkName = _
                        VTEquationCaptionBookmarkName(formulaId)
                Case Else
                    sourceBookmarkName = _
                        VTEquationSequenceNumberBookmarkName(formulaId)
            End Select
            If documentObject.Bookmarks.Exists(sourceBookmarkName) Then
                expectedResult = documentObject.Bookmarks( _
                    sourceBookmarkName).Range.Text
                expectedResult = Replace$(expectedResult, vbCr, "")
                expectedResult = Replace$(expectedResult, Chr$(7), "")
                expectedResult = Trim$(expectedResult)
                If StrComp(targetBookmarkName, sourceBookmarkName, _
                   vbTextCompare) <> 0 Then
                    VTReplaceBookmarkRangeIncludingHidden _
                        documentObject, targetBookmarkName, _
                        documentObject.Bookmarks( _
                            sourceBookmarkName).Range.Duplicate
                End If
                VTRepairBodyReferenceFieldsForTarget _
                    documentObject, targetBookmarkName, expectedResult
            Else
                VTReplaceBodyReferenceTargetWithDeletedMarker _
                    documentObject, targetBookmarkName
            End If
        End If
    Next itemIndex
End Sub

Private Function VTSequenceBookmarkNameForField( _
    ByVal documentObject As Document, _
    ByVal sequenceField As Field) As String

    Dim candidateBookmark As Bookmark
    Dim localRange As Range

    If documentObject Is Nothing Or sequenceField Is Nothing Then Exit Function

    ' Word exposes bookmarks contained by a Range without enumerating every
    ' bookmark in the document. The exact VT_N_ result bookmark is normally
    ' returned directly; the helper paragraph covers hosts that omit an exact
    ' result bookmark from Range.Bookmarks after a field update.
    Set localRange = sequenceField.Result.Duplicate
    For Each candidateBookmark In localRange.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX Then
            VTSequenceBookmarkNameForField = candidateBookmark.Name
            Exit Function
        End If
    Next candidateBookmark

    Set localRange = sequenceField.Result.Paragraphs(1).Range.Duplicate
    For Each candidateBookmark In localRange.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX Then
            If candidateBookmark.Range.Start <= sequenceField.Result.Start And _
               candidateBookmark.Range.End >= sequenceField.Result.End Then
                VTSequenceBookmarkNameForField = candidateBookmark.Name
                Exit Function
            End If
        End If
    Next candidateBookmark

    ' Legacy Word builds can omit the bookmark from both local collections.
    ' Preserve a correctness fallback, but the normal path remains local and
    ' avoids the former Fields x Bookmarks quadratic scan.
    For Each candidateBookmark In documentObject.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX Then
            If candidateBookmark.Range.Start <= sequenceField.Result.Start And _
               candidateBookmark.Range.End >= sequenceField.Result.End Then
                VTSequenceBookmarkNameForField = candidateBookmark.Name
                Exit Function
            End If
        End If
    Next candidateBookmark
End Function

Private Function VTNumberedDisplayTableContainsFormula( _
    ByVal layoutTable As Table, _
    ByVal formulaId As String) As Boolean

    Dim centerRange As Range

    If layoutTable Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then
        Exit Function
    End If
    If layoutTable.Rows.Count <> 1 Or layoutTable.Columns.Count <> 3 Then
        Exit Function
    End If
    Set centerRange = layoutTable.Cell(1, 2).Range.Duplicate
    VTNumberedDisplayTableContainsFormula = _
        (centerRange.InlineShapes.Count = 1 And _
         centerRange.OMaths.Count = 0) Or _
        (centerRange.InlineShapes.Count = 0 And _
         centerRange.OMaths.Count = 1)
End Function

Private Function VTFormulaIdForValidNumberedTable( _
    ByVal layoutTable As Table) As String

    Dim documentObject As Document
    Dim centerRange As Range
    Dim candidateBookmark As Bookmark
    Dim candidateMath As OMath
    Dim formulaShape As InlineShape
    Dim numberBookmarkName As String
    Dim suffixText As String
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim matchCount As Long

    If layoutTable Is Nothing Then Exit Function
    Set documentObject = layoutTable.Range.Document
    Set centerRange = layoutTable.Cell(1, 2).Range.Duplicate
    If centerRange.InlineShapes.Count + centerRange.OMaths.Count <> 1 Then
        Exit Function
    End If

    numberBookmarkName = VTNumberBookmarkNameForTable(layoutTable)
    If Len(numberBookmarkName) > 0 Then
        suffixText = Mid$(numberBookmarkName, _
            Len(VT_WORD_NUMBER_BOOKMARK_PREFIX) + 1)
        formulaId = VTFormulaIdFromBookmarkSuffix(suffixText)
    End If

    If Len(formulaId) = 0 And centerRange.InlineShapes.Count = 1 Then
        Set formulaShape = centerRange.InlineShapes(1)
        If VTTryParseFormulaReference( _
           formulaShape.Title, formulaId, displayMode, numbered) Then
            If displayMode <> "block" Or Not numbered Then formulaId = ""
        End If
    End If

    If Len(formulaId) = 0 And centerRange.OMaths.Count = 1 Then
        For Each candidateBookmark In documentObject.Bookmarks
            If VTTryFormulaIdFromNativeBookmark( _
               candidateBookmark.Name, suffixText) Then
                Set candidateMath = VTNativeMathForBookmark(candidateBookmark)
                If Not candidateMath Is Nothing Then
                    If candidateMath.Range.Start >= centerRange.Start And _
                       candidateMath.Range.End <= centerRange.End Then
                        matchCount = matchCount + 1
                        formulaId = suffixText
                    End If
                End If
            End If
        Next candidateBookmark
        If matchCount <> 1 Then formulaId = ""
    End If

    If Len(formulaId) > 0 And _
       VTNumberedDisplayTableContainsFormula(layoutTable, formulaId) Then
        VTFormulaIdForValidNumberedTable = formulaId
    End If
End Function

Private Function VTNumberedFormulaRangeForId( _
    ByVal documentObject As Document, _
    ByVal formulaId As String) As Range

    Dim numberRange As Range
    Dim formulaContainer As Range
    Dim repairedNumberRange As Range
    Dim nativeRange As Range
    Dim layoutTable As Table
    Dim nativeMath As OMath
    Dim visibleNumberField As Field
    Dim numberBookmarkName As String
    Dim nativeBookmarkName As String

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then
        Exit Function
    End If
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(numberBookmarkName) Then
        Set numberRange = documentObject.Bookmarks( _
            numberBookmarkName).Range.Duplicate

        If numberRange.Information(wdWithInTable) Then
            Set layoutTable = numberRange.Tables(1)
            If layoutTable.Rows.Count = 1 And _
               layoutTable.Columns.Count = 3 Then
                Set formulaContainer = _
                    layoutTable.Cell(1, 2).Range.Duplicate
            End If
        Else
            Set formulaContainer = _
                numberRange.Paragraphs(1).Range.Duplicate
        End If

        If Not formulaContainer Is Nothing Then
            If formulaContainer.InlineShapes.Count = 1 And _
               formulaContainer.OMaths.Count = 0 Then
                Set VTNumberedFormulaRangeForId = _
                    formulaContainer.InlineShapes(1).Range.Duplicate
                Exit Function
            End If
            If formulaContainer.InlineShapes.Count = 0 And _
               formulaContainer.OMaths.Count = 1 Then
                Set VTNumberedFormulaRangeForId = _
                    formulaContainer.OMaths(1).Range.Duplicate
                Exit Function
            End If
        End If
    End If

    ' Word can expand VT_R_ across a newly inserted blank paragraph at the start
    ' of a built-up OMath paragraph. Recover only through this formula's exact
    ' VT_O_ identity; never scan for the nearest Equation, which could select a
    ' neighboring formula and reproduce the destructive insertion bug.
    nativeBookmarkName = VTNativeFormulaBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(nativeBookmarkName) Then
        Exit Function
    End If
    Set nativeMath = VTNativeMathForBookmark( _
        documentObject.Bookmarks(nativeBookmarkName))
    If nativeMath Is Nothing Then Exit Function
    Set nativeRange = nativeMath.Range.Duplicate
    If nativeRange.Information(wdWithInTable) Then Exit Function

    Set visibleNumberField = VTNativeEquationArrayReferenceField( _
        nativeRange, formulaId)
    If visibleNumberField Is Nothing Then Exit Function
    If Not VTNativeEquationNumberIsInsideMath( _
       nativeRange, visibleNumberField) Then Exit Function
    Set repairedNumberRange = VTNativeEquationArrayNumberRange( _
        nativeRange, visibleNumberField)
    If repairedNumberRange Is Nothing Then Exit Function

    VTSetEquationNumberBookmarkExact _
        documentObject, formulaId, repairedNumberRange
    VTSetNativeFormulaBookmark documentObject, nativeRange, formulaId
    Set VTNumberedFormulaRangeForId = nativeRange.Duplicate
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

Private Function VTValidNumberedFormulaIds( _
    ByVal documentObject As Document) As Variant

    Dim candidateField As Field
    Dim visibleNumberField As Field
    Dim layoutTable As Table
    Dim numberRange As Range
    Dim sequenceParagraph As Range
    Dim formulaRange As Range
    Dim formulaParagraph As Range
    Dim seenIds As New Collection
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim captionBookmarkName As String
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim formulaIsLive As Boolean
    Dim ids() As String
    Dim itemCount As Long

    If documentObject Is Nothing Then Exit Function

    ' Every true SEQ field remains in document order. Image formulas keep SEQ
    ' in their visible paragraph; native OMath formulas keep SEQ in the compact
    ' paragraph immediately after the formula and expose only a REF inside #().
    For Each candidateField In documentObject.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            sequenceBookmarkName = VTSequenceBookmarkNameForField( _
                documentObject, candidateField)
            formulaId = VTFormulaIdFromSequenceBookmarkName( _
                sequenceBookmarkName)
            formulaIsLive = False
            If Len(formulaId) > 0 Then
                numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
                captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
                If documentObject.Bookmarks.Exists(numberBookmarkName) And _
                   documentObject.Bookmarks.Exists(captionBookmarkName) Then
                    Set numberRange = documentObject.Bookmarks( _
                        numberBookmarkName).Range.Duplicate
                    If numberRange.Information(wdWithInTable) Then
                        Set layoutTable = numberRange.Tables(1)
                        formulaIsLive = _
                            VTNumberedDisplayTableContainsFormula( _
                                layoutTable, formulaId)
                    Else
                        Set formulaRange = VTNumberedFormulaRangeForId( _
                            documentObject, formulaId)
                        If Not formulaRange Is Nothing Then
                            Set formulaParagraph = _
                                VTWordParagraphContainingFormula(formulaRange)
                            Set sequenceParagraph = _
                                candidateField.Result.Paragraphs(1).Range.Duplicate
                            If Not formulaParagraph Is Nothing And _
                               VTTryReadWordFormulaFormat( _
                                   documentObject, formulaId, _
                                   displayMode, numbered) Then
                                If displayMode = "block" And numbered Then
                                    If formulaRange.InlineShapes.Count = 1 And _
                                       formulaRange.OMaths.Count = 0 Then
                                        formulaIsLive = _
                                            VTHelperParagraphOwnsNativeEquationSequence( _
                                                sequenceParagraph) And _
                                            sequenceParagraph.Start >= _
                                                formulaParagraph.End
                                        If formulaIsLive Then
                                            Set visibleNumberField = _
                                                VTImageEquationReferenceField( _
                                                    formulaRange, formulaId)
                                            formulaIsLive = _
                                                Not visibleNumberField Is Nothing
                                        End If
                                        If formulaIsLive Then
                                            Set numberRange = _
                                                VTImageEquationNumberRange( _
                                                    formulaRange, _
                                                    visibleNumberField)
                                            formulaIsLive = _
                                                Not numberRange Is Nothing
                                        End If
                                    ElseIf formulaRange.InlineShapes.Count = 0 And _
                                           formulaRange.OMaths.Count = 1 And _
                                           VTHelperParagraphOwnsNativeEquationSequence( _
                                               sequenceParagraph) And _
                                           sequenceParagraph.Start >= _
                                               formulaParagraph.End Then
                                        Set visibleNumberField = _
                                            VTNativeEquationArrayReferenceField( _
                                                formulaRange, formulaId)
                                        formulaIsLive = _
                                            Not visibleNumberField Is Nothing
                                        If formulaIsLive Then
                                            formulaIsLive = _
                                                VTNativeEquationNumberIsInsideMath( _
                                                    formulaRange, _
                                                    visibleNumberField)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If formulaIsLive Then
                If VTCollectionContainsText(seenIds, formulaId) Then
                    Err.Raise vbObjectError + 7553, "VisualTeX", _
                        "Two numbered VisualTeX formulas use the same formula id" & _
                        " [formulaId=" & formulaId & "]."
                End If
                VTCollectionAddUniqueText seenIds, formulaId
                itemCount = itemCount + 1
                ReDim Preserve ids(1 To itemCount)
                ids(itemCount) = formulaId
            End If
        End If
    Next candidateField
    If itemCount > 0 Then VTValidNumberedFormulaIds = ids
End Function

Private Function VTVariantArrayCount(ByVal values As Variant) As Long
    On Error GoTo NoValues
    If IsArray(values) Then
        VTVariantArrayCount = UBound(values) - LBound(values) + 1
    End If
    Exit Function
NoValues:
    VTVariantArrayCount = 0
End Function

Private Sub VTReplaceDeletedEquationBodyReferences( _
    ByVal documentObject As Document, _
    ByVal sequenceBookmarkName As String)

    Dim candidateField As Field
    Dim replacementRange As Range
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim fieldIndex As Long

    If documentObject Is Nothing Or Len(sequenceBookmarkName) = 0 Then Exit Sub
    For fieldIndex = documentObject.Fields.Count To 1 Step -1
        Set candidateField = documentObject.Fields(fieldIndex)
        If candidateField.Type = wdFieldRef And _
           InStr(1, candidateField.Code.Text, sequenceBookmarkName, _
           vbTextCompare) > 0 And _
           Not candidateField.Result.Information(wdWithInTable) Then
            fieldStart = VTEquationFieldStart(candidateField)
            fieldEnd = VTEquationFieldEnd(candidateField)
            If fieldStart > 0 Then
                If documentObject.Range( _
                   Start:=fieldStart - 1, End:=fieldStart).Text = "(" Then
                    fieldStart = fieldStart - 1
                End If
            End If
            If fieldEnd < documentObject.Content.End - 1 Then
                If documentObject.Range( _
                   Start:=fieldEnd, End:=fieldEnd + 1).Text = ")" Then
                    fieldEnd = fieldEnd + 1
                End If
            End If
            Set replacementRange = documentObject.Range( _
                Start:=fieldStart, End:=fieldEnd)
            replacementRange.Text = "(deleted equation)"
            With replacementRange.Font
                .Hidden = False
                .Color = wdColorAutomatic
                .Position = 0
                .Size = VTVisibleEquationNumberFontSize(documentObject)
            End With
        End If
    Next fieldIndex
End Sub

Private Function VTHelperParagraphOwnsNativeEquationSequence( _
    ByVal paragraphRange As Range) As Boolean

    Dim candidateField As Field

    If paragraphRange Is Nothing Or _
       paragraphRange.Information(wdWithInTable) Or _
       paragraphRange.Fields.Count <> 1 Or _
       paragraphRange.InlineShapes.Count <> 0 Or _
       paragraphRange.OMaths.Count <> 0 Then Exit Function
    Set candidateField = paragraphRange.Fields(1)
    VTHelperParagraphOwnsNativeEquationSequence = _
        VTIsNativeEquationSequenceField( _
            candidateField, VTNativeEquationLabelName())
End Function

Private Function VTIsDetachedVisualTeXNativeSequenceHelper( _
    ByVal paragraphRange As Range) As Boolean

    Dim candidateField As Field
    Dim contentRange As Range
    Dim paragraphText As String
    Dim resultText As String

    If paragraphRange Is Nothing Or _
       Not VTHelperParagraphOwnsNativeEquationSequence( _
           paragraphRange) Then Exit Function
    Set candidateField = paragraphRange.Fields(1)

    ' A live VisualTeX helper always owns a VT_N_ Bookmark. Once the user
    ' deletes the numbered OMML formula, Word can remove the formula Bookmarks
    ' while leaving this compact SEQ paragraph behind. Never classify a live
    ' helper or an ordinary Word caption as detached.
    If Len(VTSequenceBookmarkNameForField( _
       paragraphRange.Document, candidateField)) > 0 Then Exit Function

    Set contentRange = paragraphRange.Duplicate
    If contentRange.End > contentRange.Start Then
        contentRange.End = contentRange.End - 1
    End If
    paragraphText = Trim$(Replace$(Replace$( _
        contentRange.Text, vbTab, ""), ChrW(160), " "))
    resultText = Trim$(Replace$(Replace$( _
        candidateField.Result.Text, vbTab, ""), ChrW(160), " "))
    If paragraphText <> resultText Then Exit Function

    ' This exact layout signature is created only by
    ' VTFormatHiddenEquationParagraph. It distinguishes a deleted VisualTeX
    ' OMML helper from a user-created native Word Equation caption.
    With paragraphRange.ParagraphFormat
        If .Alignment <> wdAlignParagraphLeft Or _
           Abs(.LeftIndent + 360!) > 0.2 Or _
           Abs(.RightIndent) > 0.2 Or _
           Abs(.FirstLineIndent) > 0.2 Or _
           Abs(.SpaceBefore) > 0.2 Or _
           Abs(.SpaceAfter) > 0.2 Or _
           .LineSpacingRule <> wdLineSpaceExactly Or _
           Abs(.LineSpacing - 1!) > 0.2 Then Exit Function
    End With

    VTIsDetachedVisualTeXNativeSequenceHelper = True
End Function

Private Function VTDetachedVisualTeXNativeSequenceHelperNearRange( _
    ByVal selectedRange As Range) As Range

    Dim documentObject As Document
    Dim currentParagraph As Range
    Dim probeParagraph As Range

    If selectedRange Is Nothing Then Exit Function
    Set documentObject = selectedRange.Document
    Set currentParagraph = selectedRange.Paragraphs(1).Range.Duplicate

    If VTIsDetachedVisualTeXNativeSequenceHelper(currentParagraph) Then
        Set VTDetachedVisualTeXNativeSequenceHelperNearRange = _
            currentParagraph
        Exit Function
    End If

    If currentParagraph.Start > 0 Then
        Set probeParagraph = documentObject.Range( _
            Start:=currentParagraph.Start - 1, _
            End:=currentParagraph.Start - 1).Paragraphs(1).Range.Duplicate
        If VTIsDetachedVisualTeXNativeSequenceHelper( _
           probeParagraph) Then
            Set VTDetachedVisualTeXNativeSequenceHelperNearRange = _
                probeParagraph
            Exit Function
        End If
    End If

    If currentParagraph.End < documentObject.Content.End Then
        Set probeParagraph = documentObject.Range( _
            Start:=currentParagraph.End, _
            End:=currentParagraph.End).Paragraphs(1).Range.Duplicate
        If VTIsDetachedVisualTeXNativeSequenceHelper( _
           probeParagraph) Then
            Set VTDetachedVisualTeXNativeSequenceHelperNearRange = _
                probeParagraph
        End If
    End If
End Function

Private Function VTPruneDetachedVisualTeXNativeSequenceHelpers( _
    ByVal documentObject As Document) As Long

    Dim candidateField As Field
    Dim helperParagraph As Range
    Dim helperStarts() As Long
    Dim helperCount As Long
    Dim itemIndex As Long

    If documentObject Is Nothing Then Exit Function

    ' First capture stable paragraph anchors without mutating Fields. Word for
    ' Mac can throw 5941 or skip later fields when a For Each Fields traversal
    ' deletes or rebuilds a field in the same collection.
    For Each candidateField In documentObject.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            Set helperParagraph = _
                candidateField.Result.Paragraphs(1).Range.Duplicate
            If VTIsDetachedVisualTeXNativeSequenceHelper( _
               helperParagraph) Then
                helperCount = helperCount + 1
                If helperCount = 1 Then
                    ReDim helperStarts(1 To 1)
                Else
                    ReDim Preserve helperStarts(1 To helperCount)
                End If
                helperStarts(helperCount) = helperParagraph.Start
            End If
        End If
    Next candidateField

    ' Delete from the end so earlier anchors remain stable. Revalidate every
    ' candidate immediately before deletion in case Word changed the document
    ' between the scan and cleanup.
    For itemIndex = helperCount To 1 Step -1
        If helperStarts(itemIndex) < documentObject.Content.End Then
            Set helperParagraph = documentObject.Range( _
                Start:=helperStarts(itemIndex), _
                End:=helperStarts(itemIndex)).Paragraphs(1).Range.Duplicate
            If VTIsDetachedVisualTeXNativeSequenceHelper( _
               helperParagraph) Then
                helperParagraph.Delete
                VTPruneDetachedVisualTeXNativeSequenceHelpers = _
                    VTPruneDetachedVisualTeXNativeSequenceHelpers + 1
            End If
        End If
    Next itemIndex
End Function

Private Sub VTDeleteEquationNumberScaffold( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    Optional ByVal deleteTable As Boolean = True)

    Dim numberBookmarkName As String
    Dim sequenceBookmarkName As String
    Dim captionBookmarkName As String
    Dim nativeBookmarkName As String
    Dim sequenceHelperParagraph As Range
    Dim formulaParagraph As Range
    Dim numberRange As Range
    Dim paragraphScaffold As Range
    Dim paragraphContent As Range
    Dim layoutTable As Table
    Dim sequenceField As Field
    Dim paragraphNumber As Boolean

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then
        Exit Sub
    End If
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    sequenceBookmarkName = VTEquationSequenceNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
    nativeBookmarkName = VTNativeFormulaBookmarkName(formulaId)

    If documentObject.Bookmarks.Exists(numberBookmarkName) Then
        Set numberRange = documentObject.Bookmarks( _
            numberBookmarkName).Range.Duplicate
        On Error Resume Next
        If numberRange.Information(wdWithInTable) Then
            Set layoutTable = numberRange.Tables(1)
        Else
            paragraphNumber = True
        End If
        On Error GoTo 0
    End If
    If documentObject.Bookmarks.Exists(captionBookmarkName) Then
        Set sequenceHelperParagraph = documentObject.Bookmarks( _
            captionBookmarkName).Range.Paragraphs(1).Range.Duplicate
    ElseIf documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        Set sequenceField = VTEquationSequenceFieldForBookmark( _
            documentObject, sequenceBookmarkName)
        If Not sequenceField Is Nothing Then
            Set sequenceHelperParagraph = _
                sequenceField.Result.Paragraphs(1).Range.Duplicate
        End If
    End If

    VTReplaceDeletedEquationBodyReferences _
        documentObject, sequenceBookmarkName
    If documentObject.Bookmarks.Exists(numberBookmarkName) Then
        documentObject.Bookmarks(numberBookmarkName).Delete
    End If
    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        documentObject.Bookmarks(sequenceBookmarkName).Delete
    End If
    If documentObject.Bookmarks.Exists(captionBookmarkName) Then
        documentObject.Bookmarks(captionBookmarkName).Delete
    End If
    If documentObject.Bookmarks.Exists(nativeBookmarkName) Then
        documentObject.Bookmarks(nativeBookmarkName).Delete
    End If

    If paragraphNumber And Not numberRange Is Nothing Then
        Set paragraphScaffold = numberRange.Duplicate
        If paragraphScaffold.Start > 0 Then
            If documentObject.Range( _
               Start:=paragraphScaffold.Start - 1, _
               End:=paragraphScaffold.Start).Text = vbTab Then
                paragraphScaffold.Start = paragraphScaffold.Start - 1
            End If
        End If
        Set formulaParagraph = _
            paragraphScaffold.Paragraphs(1).Range.Duplicate
        paragraphScaffold.Delete

        ' If the formula itself has already been deleted, clear any remaining
        ' visible number tail but preserve the ordinary paragraph mark.
        If formulaParagraph.InlineShapes.Count = 0 And _
           formulaParagraph.OMaths.Count = 0 Then
            Set paragraphContent = formulaParagraph.Duplicate
            If paragraphContent.End > paragraphContent.Start Then
                paragraphContent.End = paragraphContent.End - 1
            End If
            If Not VTWordRangeHasMeaningfulText(paragraphContent) Then
                paragraphContent.Text = ""
                VTNormalizePlainWordParagraph formulaParagraph
            End If
        End If
    End If
    If Not sequenceHelperParagraph Is Nothing Then
        If VTHelperParagraphOwnsNativeEquationSequence( _
           sequenceHelperParagraph) Then
            sequenceHelperParagraph.Delete
        End If
    End If
    If deleteTable And Not layoutTable Is Nothing Then layoutTable.Delete

    VTDeleteWordLatexPayload documentObject, formulaId
    VTDeleteWordOmmlPayload documentObject, formulaId
    VTDeleteWordMetadataPayload documentObject, formulaId
    VTDeleteDocumentVariable _
        documentObject, VTWordFormatVariableName(formulaId)
End Sub

Private Function VTNumberedTableScaffoldComplete( _
    ByVal layoutTable As Table, _
    ByVal formulaId As String) As Boolean

    Dim documentObject As Document
    Dim candidateField As Field
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim captionBookmarkName As String

    If layoutTable Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then
        Exit Function
    End If
    Set documentObject = layoutTable.Range.Document
    sequenceBookmarkName = VTEquationSequenceNumberBookmarkName(formulaId)
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(sequenceBookmarkName) Or _
       Not documentObject.Bookmarks.Exists(numberBookmarkName) Or _
       Not documentObject.Bookmarks.Exists(captionBookmarkName) Then
        Exit Function
    End If
    For Each candidateField In layoutTable.Cell(1, 3).Range.Fields
        If candidateField.Type = wdFieldRef And _
           InStr(1, candidateField.Code.Text, sequenceBookmarkName, _
           vbTextCompare) > 0 Then
            VTNumberedTableScaffoldComplete = True
            Exit Function
        End If
    Next candidateField
End Function

Private Sub VTRepairLiveNumberedTableScaffolds( _
    ByVal documentObject As Document)

    Dim layoutTable As Table
    Dim formulaId As String
    Dim tableIndex As Long

    If documentObject Is Nothing Then Exit Sub
    For tableIndex = 1 To documentObject.Tables.Count
        Set layoutTable = documentObject.Tables(tableIndex)
        formulaId = VTFormulaIdForValidNumberedTable(layoutTable)
        If Len(formulaId) > 0 Then
            If Not VTNumberedTableScaffoldComplete( _
               layoutTable, formulaId) Then
                VTEnsureEquationNumberFields layoutTable, formulaId
            End If
        End If
    Next tableIndex
End Sub

Private Sub VTPruneUnbookmarkedEmptyNumberTables( _
    ByVal documentObject As Document)

    Dim layoutTable As Table
    Dim centerRange As Range
    Dim candidateField As Field
    Dim sequenceBookmarkName As String
    Dim suffixText As String
    Dim formulaId As String
    Dim tableIndex As Long

    If documentObject Is Nothing Then Exit Sub
    For tableIndex = documentObject.Tables.Count To 1 Step -1
        If tableIndex > documentObject.Tables.Count Then GoTo NextTable
        Set layoutTable = documentObject.Tables(tableIndex)
        If layoutTable.Rows.Count <> 1 Or _
           layoutTable.Columns.Count <> 3 Then GoTo NextTable
        Set centerRange = layoutTable.Cell(1, 2).Range.Duplicate
        If centerRange.InlineShapes.Count <> 0 Or _
           centerRange.OMaths.Count <> 0 Or _
           VTWordRangeHasMeaningfulText(centerRange) Then GoTo NextTable

        sequenceBookmarkName = ""
        For Each candidateField In layoutTable.Cell(1, 3).Range.Fields
            If candidateField.Type = wdFieldRef Then
                sequenceBookmarkName = VTBookmarkNameInFieldCode( _
                    documentObject, candidateField.Code.Text, _
                    VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)
                If Len(sequenceBookmarkName) > 0 Then Exit For
            End If
        Next candidateField
        If Len(sequenceBookmarkName) = 0 Then GoTo NextTable
        suffixText = Mid$(sequenceBookmarkName, _
            Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX) + 1)
        formulaId = VTFormulaIdFromBookmarkSuffix(suffixText)
        If Len(formulaId) = 0 Then GoTo NextTable
        VTDeleteEquationNumberScaffold documentObject, formulaId, False
        layoutTable.Delete
NextTable:
    Next tableIndex
End Sub

Private Function VTMigrateLegacyImageEquationSequenceLayouts( _
    ByVal documentObject As Document) As Long

    Dim candidateField As Field
    Dim formulaRange As Range
    Dim migratedRange As Range
    Dim formulaParagraph As Range
    Dim sequenceParagraph As Range
    Dim formulaShape As InlineShape
    Dim referenceBindings As Collection
    Dim migrationIds As New Collection
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim formulaId As String
    Dim itemIndex As Long

    If documentObject Is Nothing Then Exit Function

    ' Snapshot identities before changing any fields. r29 image formulas place
    ' the true SEQ in the visible image paragraph; r30 moves that same VT_N_
    ' identity to a compact helper paragraph and replaces the visible number
    ' with a REF. VisualTeX and Word-native body references are rebound after
    ' the migration so their displayed values and targets remain stable.
    For Each candidateField In documentObject.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            sequenceBookmarkName = VTSequenceBookmarkNameForField( _
                documentObject, candidateField)
            formulaId = VTFormulaIdFromSequenceBookmarkName( _
                sequenceBookmarkName)
            If Len(formulaId) > 0 Then
                numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
                If documentObject.Bookmarks.Exists(numberBookmarkName) Then
                    Set formulaRange = VTNumberedFormulaRangeForId( _
                        documentObject, formulaId)
                    If Not formulaRange Is Nothing Then
                        If formulaRange.InlineShapes.Count = 1 And _
                           formulaRange.OMaths.Count = 0 Then
                            Set formulaParagraph = _
                                VTWordParagraphContainingFormula(formulaRange)
                            Set sequenceParagraph = _
                                candidateField.Result.Paragraphs(1).Range.Duplicate
                            If Not formulaParagraph Is Nothing Then
                                If sequenceParagraph.Start = _
                                   formulaParagraph.Start Then
                                    VTCollectionAddUniqueText _
                                        migrationIds, formulaId
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next candidateField

    If migrationIds.Count = 0 Then Exit Function
    Set referenceBindings = _
        VTCaptureBodyEquationReferenceBindings(documentObject)

    For itemIndex = 1 To migrationIds.Count
        formulaId = CStr(migrationIds(itemIndex))
        Set formulaRange = VTNumberedFormulaRangeForId( _
            documentObject, formulaId)
        If formulaRange Is Nothing Then
            Err.Raise vbObjectError + 7568, "VisualTeX", _
                "A legacy image Equation disappeared during native-reference migration."
        End If
        If formulaRange.InlineShapes.Count <> 1 Or _
           formulaRange.OMaths.Count <> 0 Then
            Err.Raise vbObjectError + 7568, "VisualTeX", _
                "A legacy image Equation became ambiguous during native-reference migration."
        End If
        Set formulaShape = formulaRange.InlineShapes(1)
        Set migratedRange = VTInsertEquationNumber( _
            formulaShape, formulaId, "VisualTeX formula", True)
        VTMigrateLegacyImageEquationSequenceLayouts = _
            VTMigrateLegacyImageEquationSequenceLayouts + 1
    Next itemIndex

    VTReconcileEquationNumbers documentObject
    VTRestoreBodyEquationReferenceBindings _
        documentObject, referenceBindings
End Function

Private Function VTPruneOrphanedEquationNumberScaffolds( _
    ByVal documentObject As Document) As Long

    Dim candidateBookmark As Bookmark
    Dim sequenceField As Field
    Dim helperParagraph As Range
    Dim orphanIds As New Collection
    Dim validIds As New Collection
    Dim validFormulaIds As Variant
    Dim formulaId As String
    Dim suffixText As String
    Dim sequenceBookmarkName As String
    Dim fieldIndex As Long
    Dim itemIndex As Long

    If documentObject Is Nothing Then Exit Function

    VTMigrateLegacyImageEquationSequenceLayouts documentObject
    VTPruneUnbookmarkedEmptyNumberTables documentObject
    VTPruneDetachedVisualTeXNativeSequenceHelpers documentObject
    validFormulaIds = VTValidNumberedFormulaIds(documentObject)
    If IsArray(validFormulaIds) Then
        For itemIndex = LBound(validFormulaIds) To UBound(validFormulaIds)
            VTCollectionAddUniqueText validIds, CStr(validFormulaIds(itemIndex))
        Next itemIndex
    End If

    For Each candidateBookmark In documentObject.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_NUMBER_BOOKMARK_PREFIX Then
            suffixText = Mid$(candidateBookmark.Name, _
                Len(VT_WORD_NUMBER_BOOKMARK_PREFIX) + 1)
            formulaId = VTFormulaIdFromBookmarkSuffix(suffixText)
            If Len(formulaId) > 0 And _
               Not VTCollectionContainsText(validIds, formulaId) Then
                VTCollectionAddUniqueText orphanIds, formulaId
            End If
        End If
    Next candidateBookmark

    For itemIndex = 1 To orphanIds.Count
        VTDeleteEquationNumberScaffold _
            documentObject, CStr(orphanIds(itemIndex)), True
    Next itemIndex

    validFormulaIds = VTValidNumberedFormulaIds(documentObject)
    Set validIds = New Collection
    If IsArray(validFormulaIds) Then
        For itemIndex = LBound(validFormulaIds) To UBound(validFormulaIds)
            VTCollectionAddUniqueText validIds, CStr(validFormulaIds(itemIndex))
        Next itemIndex
    End If

    For fieldIndex = documentObject.Fields.Count To 1 Step -1
        If fieldIndex > documentObject.Fields.Count Then GoTo NextSequenceField
        Set sequenceField = documentObject.Fields(fieldIndex)
        If VTIsNativeEquationSequenceField( _
           sequenceField, VTNativeEquationLabelName()) Then
            sequenceBookmarkName = VTSequenceBookmarkNameForField( _
                documentObject, sequenceField)
            formulaId = ""
            If Len(sequenceBookmarkName) > 0 Then
                suffixText = Mid$(sequenceBookmarkName, _
                    Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX) + 1)
                formulaId = VTFormulaIdFromBookmarkSuffix(suffixText)
            End If
            ' Never delete an unbookmarked native Equation caption: it can be
            ' ordinary user content created by Word. Only VisualTeX captions
            ' with a VT_N_ identity participate in VisualTeX garbage collection.
            If Len(formulaId) > 0 And _
               Not VTCollectionContainsText(validIds, formulaId) Then
                VTDeleteEquationNumberScaffold _
                    documentObject, formulaId, False
            End If
        End If
NextSequenceField:
    Next fieldIndex

    VTRepairLiveNumberedTableScaffolds documentObject
    validFormulaIds = VTValidNumberedFormulaIds(documentObject)
    VTPruneOrphanedEquationNumberScaffolds = _
        VTVariantArrayCount(validFormulaIds)
End Function

Public Sub VTCleanupOrphanedNumberedDisplaySelection( _
    ByVal selectedRange As Range)

    Dim documentObject As Document
    Dim layoutTable As Table
    Dim centerContent As Range
    Dim caretRange As Range
    Dim paragraphRange As Range
    Dim detachedHelper As Range
    Dim candidateBookmark As Bookmark
    Dim numberBookmarkName As String
    Dim suffixText As String
    Dim formulaId As String
    Dim cleanupStart As Long

    If selectedRange Is Nothing Or _
       VTWordInternalMutationActive() Then Exit Sub

    ' Legacy numbered tables keep their established cleanup path.
    Set layoutTable = VTNumberedDisplayTableNearRange(selectedRange)
    If Not layoutTable Is Nothing Then
        If layoutTable.Cell(1, 2).Range.InlineShapes.Count <> 0 Or _
           layoutTable.Cell(1, 2).Range.OMaths.Count <> 0 Then Exit Sub
        Set centerContent = layoutTable.Cell(1, 2).Range.Duplicate
        If centerContent.End > centerContent.Start Then
            centerContent.End = centerContent.End - 1
        End If
        If VTWordRangeHasMeaningfulText(centerContent) Then Exit Sub

        numberBookmarkName = VTNumberBookmarkNameForTable(layoutTable)
        If Len(numberBookmarkName) = 0 Then Exit Sub
        suffixText = Mid$(numberBookmarkName, _
            Len(VT_WORD_NUMBER_BOOKMARK_PREFIX) + 1)
        formulaId = VTFormulaIdFromBookmarkSuffix(suffixText)
        If Len(formulaId) = 0 Then Exit Sub
        Set documentObject = layoutTable.Range.Document
        cleanupStart = layoutTable.Range.Start
        VTDeleteEquationNumberScaffold documentObject, formulaId, True
        GoTo CleanupFinished
    End If

    ' A numbered native OMML formula stores its SEQ in a compact helper
    ' paragraph immediately after the formula. Deleting the OMath can remove
    ' every VT_* Bookmark before the watcher runs, so inspect only the selected
    ' paragraph and its immediate neighbors for the exact VisualTeX helper
    ' signature. Ordinary Word captions do not use this layout.
    Set detachedHelper = _
        VTDetachedVisualTeXNativeSequenceHelperNearRange(selectedRange)
    If Not detachedHelper Is Nothing Then
        Set documentObject = detachedHelper.Document
        cleanupStart = detachedHelper.Start
        detachedHelper.Delete
        GoTo CleanupFinished
    End If

    ' Numbered image formulas and legacy same-paragraph layouts retain a local
    ' VT_R_ Bookmark after the image or OMath is removed.
    Set paragraphRange = selectedRange.Paragraphs(1).Range.Duplicate
    If paragraphRange.Information(wdWithInTable) Or _
       paragraphRange.InlineShapes.Count <> 0 Or _
       paragraphRange.OMaths.Count <> 0 Then Exit Sub
    For Each candidateBookmark In paragraphRange.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_NUMBER_BOOKMARK_PREFIX Then
            numberBookmarkName = candidateBookmark.Name
            Exit For
        End If
    Next candidateBookmark
    If Len(numberBookmarkName) = 0 Then Exit Sub
    suffixText = Mid$(numberBookmarkName, _
        Len(VT_WORD_NUMBER_BOOKMARK_PREFIX) + 1)
    formulaId = VTFormulaIdFromBookmarkSuffix(suffixText)
    If Len(formulaId) = 0 Then Exit Sub
    Set documentObject = paragraphRange.Document
    cleanupStart = paragraphRange.Start
    VTDeleteEquationNumberScaffold documentObject, formulaId, False

CleanupFinished:
    VTReconcileEquationNumbers documentObject
    On Error Resume Next
    If cleanupStart > documentObject.Content.End - 1 Then
        cleanupStart = documentObject.Content.End - 1
    End If
    If cleanupStart < 0 Then cleanupStart = 0
    Set caretRange = documentObject.Range( _
        Start:=cleanupStart, End:=cleanupStart)
    Set paragraphRange = caretRange.Paragraphs(1).Range.Duplicate
    VTNormalizePlainWordParagraph paragraphRange
    If documentObject Is ActiveDocument Then caretRange.Select
    On Error GoTo 0
End Sub

Private Function VTAnyOpenDocumentHasEquationNumbers() As Boolean
    Dim documentObject As Document
    Dim candidateBookmark As Bookmark
    Dim candidateField As Field
    Dim helperParagraph As Range

    For Each documentObject In Documents
        For Each candidateBookmark In documentObject.Bookmarks
            If Left$(candidateBookmark.Name, _
               Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
               VT_WORD_NUMBER_BOOKMARK_PREFIX Then
                VTAnyOpenDocumentHasEquationNumbers = True
                Exit Function
            End If
        Next candidateBookmark

        ' Keep the watcher alive for one cleanup cycle when the last numbered
        ' OMML formula was deleted but its compact VisualTeX SEQ helper remains.
        For Each candidateField In documentObject.Fields
            If VTIsNativeEquationSequenceField( _
               candidateField, VTNativeEquationLabelName()) Then
                Set helperParagraph = _
                    candidateField.Result.Paragraphs(1).Range.Duplicate
                If VTIsDetachedVisualTeXNativeSequenceHelper( _
                   helperParagraph) Then
                    VTAnyOpenDocumentHasEquationNumbers = True
                    Exit Function
                End If
            End If
        Next candidateField
    Next documentObject
End Function

Private Sub VTNormalizeSelectedEquationReferences()
    Dim scanRange As Range

    If Documents.Count = 0 Then Exit Sub
    Set scanRange = Selection.Range.Duplicate
    If scanRange.Start = scanRange.End Then
        Set scanRange = scanRange.Paragraphs(1).Range.Duplicate
    End If
    If scanRange.Fields.Count = 0 Then Exit Sub

    ' Repair manual F9/context-menu updates only where the user is working.
    ' A collapsed caret scans one paragraph; an explicit multi-paragraph
    ' selection scans that selection. Idle typing no longer enumerates every
    ' REF field in a document once per second.
    VTNormalizeBodyEquationReferenceVisibilityInRange _
        scanRange.Document, scanRange
End Sub

Private Sub VTEnsureOrphanWatchScheduled()
    On Error GoTo ScheduleFailed
    If VT_WORD_ORPHAN_WATCH_SCHEDULED Or _
       VT_WORD_ORPHAN_WATCH_RUNNING Then Exit Sub
    If Documents.Count = 0 Then Exit Sub
    If Not VTAnyOpenDocumentHasEquationNumbers() Then Exit Sub

    VT_WORD_ORPHAN_WATCH_SCHEDULED = True
    Application.OnTime _
        When:=Now + TimeSerial(0, 0, 1), _
        name:="VisualTeX_WatchOrphanedNumberedDisplay"
    Exit Sub

ScheduleFailed:
    VT_WORD_ORPHAN_WATCH_SCHEDULED = False
End Sub

Public Sub VisualTeX_WatchOrphanedNumberedDisplay()
    VT_WORD_ORPHAN_WATCH_SCHEDULED = False
    If VT_WORD_ORPHAN_WATCH_RUNNING Then Exit Sub

    VT_WORD_ORPHAN_WATCH_RUNNING = True
    On Error Resume Next
    If Documents.Count > 0 And _
       Not VTWordInternalMutationActive() Then
        ' The selected scaffold cleanup performs a full reconciliation only
        ' when it actually removes a numbered formula. Reference visibility is
        ' repaired only in the selected range or current paragraph.
        VTCleanupOrphanedNumberedDisplaySelection Selection.Range
        VTNormalizeSelectedEquationReferences
    End If
    On Error GoTo 0
    VT_WORD_ORPHAN_WATCH_RUNNING = False
    VTEnsureOrphanWatchScheduled
End Sub

Public Sub VisualTeX_CreateInline()
    VTWordCreate "inline", False
End Sub

Public Sub VisualTeX_CreateDisplay()
    VTWordCreate "block", False
End Sub

Public Sub VisualTeX_CreateNativeInline()
    VTWordCreate "inline", False, True
End Sub

Public Sub VisualTeX_CreateNativeDisplay()
    VTWordCreate "block", False, True
End Sub

Public Sub VisualTeX_CreateNumberedDisplay()
    VTWordCreate "block", True
End Sub

' Keep every RibbonX callback resolvable from the single adapter module that is
' replaced during the validated manual VBE workflow. The legacy callback module
' did not expose these controls, so Word displayed the buttons but silently
' ignored their onAction callbacks. The onLoad callback also initializes the
' application event sink when the isolated template is attached to a test file.
Public Sub VTWordRibbonOnLoad(ByVal ribbon As IRibbonUI)
    VTInitializeWordEvents
    VTEnsureOrphanWatchScheduled
    VTRefreshWordHealthQuietly
End Sub

Public Sub VTWordRibbonNativeInline(ByVal control As IRibbonControl)
    VisualTeX_CreateNativeInline
End Sub

Public Sub VTWordRibbonNativeDisplay(ByVal control As IRibbonControl)
    VisualTeX_CreateNativeDisplay
End Sub

Public Sub VTWordRibbonCrossReference(ByVal control As IRibbonControl)
    VisualTeX_OpenEquationCrossReference
End Sub

Public Sub VisualTeX_EditSelected()
    On Error GoTo Failed

    VTRequireWritableWordDocument
    If Selection.InlineShapes.Count = 1 And VTIsVisualTeXInlineShape(Selection.InlineShapes(1)) Then
        VTWordEditInlineShape Selection.InlineShapes(1)
    Else
        VTWordEditNativeBookmark VTFindSelectedNativeFormulaBookmark(Selection)
    End If
    Exit Sub

Failed:
    VTShowError "Word edit", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_DoubleClickEditSelected()
    Dim selectedShape As InlineShape
    Dim nativeBookmark As Bookmark

    ' This entry point is invoked for every global Word double-click. It must be
    ' a strict no-op unless the clicked target is a validated VisualTeX image or
    ' native OMML bookmark. In particular, ordinary text and blank document
    ' space must never surface a VBA runtime error.
    On Error GoTo IgnoreDoubleClick
    If Documents.Count = 0 Then Exit Sub
    Set selectedShape = VTVisualTeXInlineShapeAtSelection(Selection)
    If Not selectedShape Is Nothing Then
        VTRequireWritableWordDocument
        VTRefreshWordHealthQuietly
        VTWordEditInlineShape selectedShape
        Exit Sub
    End If
    Set nativeBookmark = VTFindNativeFormulaBookmark(Selection.Range, False)
    If nativeBookmark Is Nothing Then Exit Sub
    VTRequireWritableWordDocument
    VTRefreshWordHealthQuietly
    VTWordEditNativeBookmark nativeBookmark
    Exit Sub

IgnoreDoubleClick:
    ' Preserve Word's native double-click behavior for every non-VisualTeX
    ' target, including empty space and ordinary document text.
End Sub

Public Sub VisualTeX_EditInlineShape(ByVal selectedShape As InlineShape)
    On Error GoTo Failed
    VTRequireWritableWordDocument
    VTWordEditInlineShape selectedShape
    Exit Sub
Failed:
    VTShowError "Word edit", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_EditNativeSelection(ByVal selectedRange As Range)
    On Error GoTo Failed
    VTRequireWritableWordDocument
    VTWordEditNativeBookmark VTFindNativeFormulaBookmark(selectedRange)
    Exit Sub
Failed:
    VTShowError "Word native equation edit", Err.Number, Err.Description
End Sub

Public Function VTIsVisualTeXNativeSelection(ByVal selectedRange As Range) As Boolean
    Dim nativeBookmark As Bookmark
    On Error GoTo InvalidSelection
    Set nativeBookmark = VTFindNativeFormulaBookmark(selectedRange, False)
    If nativeBookmark Is Nothing Then
        VTIsVisualTeXNativeSelection = False
    Else
        VTIsVisualTeXNativeSelection = True
    End If
    Exit Function
InvalidSelection:
    VTIsVisualTeXNativeSelection = False
End Function

Private Function VTTryRestoreVisualTeXInlineShapeReference( _
    ByVal selectedShape As InlineShape) As Boolean

    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim candidateBookmark As Bookmark
    Dim formulaId As String
    Dim candidateFormulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim storedMetadata As String
    Dim formulaReference As String
    Dim matchCount As Long

    On Error GoTo RestoreFailed
    If selectedShape Is Nothing Then Exit Function
    If Len(selectedShape.Title) <> 0 Then Exit Function
    encodedMetadata = selectedShape.AlternativeText
    If Not VTIsEncodedMetadata(encodedMetadata) Then Exit Function

    Set documentObject = selectedShape.Range.Document
    Set paragraphRange = selectedShape.Range.Paragraphs(1).Range.Duplicate
    If paragraphRange.Information(wdWithInTable) Or _
       paragraphRange.InlineShapes.Count <> 1 Or _
       paragraphRange.OMaths.Count <> 0 Then Exit Function

    For Each candidateBookmark In documentObject.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_NUMBER_BOOKMARK_PREFIX Then
            If Not candidateBookmark.Range.Information(wdWithInTable) Then
                If candidateBookmark.Range.Paragraphs(1).Range.Start = _
                   paragraphRange.Start Then
                    candidateFormulaId = VTFormulaIdFromBookmarkSuffix( _
                        Mid$(candidateBookmark.Name, _
                            Len(VT_WORD_NUMBER_BOOKMARK_PREFIX) + 1))
                    If Len(candidateFormulaId) > 0 Then
                        matchCount = matchCount + 1
                        formulaId = candidateFormulaId
                    End If
                End If
            End If
        End If
    Next candidateBookmark
    If matchCount <> 1 Then Exit Function
    If Not VTTryReadWordFormulaFormat( _
       documentObject, formulaId, displayMode, numbered) Then Exit Function
    If displayMode <> "block" Or Not numbered Then Exit Function
    If Not VTTryReadWordMetadataPayload( _
       documentObject, formulaId, storedMetadata) Then Exit Function
    If StrComp(storedMetadata, encodedMetadata, vbBinaryCompare) <> 0 Then
        Exit Function
    End If

    formulaReference = VTFormulaReference(formulaId, displayMode, numbered)
    selectedShape.Title = formulaReference
    VTTryRestoreVisualTeXInlineShapeReference = _
        (StrComp(selectedShape.Title, formulaReference, _
            vbBinaryCompare) = 0)
    Exit Function

RestoreFailed:
    VTTryRestoreVisualTeXInlineShapeReference = False
End Function

Public Function VTIsVisualTeXInlineShape(ByVal selectedShape As InlineShape) As Boolean
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean

    On Error GoTo InvalidShape
    If selectedShape Is Nothing Then Exit Function
    If Not VTIsEncodedMetadata(selectedShape.AlternativeText) Then Exit Function
    If Not VTTryParseFormulaReference( _
       selectedShape.Title, formulaId, displayMode, numbered) Then
        If Not VTTryRestoreVisualTeXInlineShapeReference( _
           selectedShape) Then Exit Function
    End If
    VTIsVisualTeXInlineShape = True
    Exit Function
InvalidShape:
    VTIsVisualTeXInlineShape = False
End Function

Public Function VTVisualTeXInlineShapeAtSelection( _
    ByVal selected As Selection) As InlineShape

    Dim probeRange As Range
    Dim candidate As InlineShape
    Dim match As InlineShape
    Dim matchCount As Long

    If selected Is Nothing Then Exit Function
    If selected.InlineShapes.Count = 1 Then
        If VTIsVisualTeXInlineShape(selected.InlineShapes(1)) Then
            Set VTVisualTeXInlineShapeAtSelection = selected.InlineShapes(1)
            Exit Function
        End If
    End If

    Set probeRange = selected.Range.Duplicate
    On Error Resume Next
    probeRange.MoveStart Unit:=wdCharacter, Count:=-1
    probeRange.MoveEnd Unit:=wdCharacter, Count:=1
    On Error GoTo InvalidSelection
    For Each candidate In probeRange.InlineShapes
        If VTIsVisualTeXInlineShape(candidate) Then
            matchCount = matchCount + 1
            Set match = candidate
        End If
    Next candidate
    If matchCount = 1 Then Set VTVisualTeXInlineShapeAtSelection = match
    Exit Function

InvalidSelection:
    Set VTVisualTeXInlineShapeAtSelection = Nothing
End Function

Private Sub VTWordEditInlineShape( _
    ByVal selectedShape As InlineShape, _
    Optional ByVal convertToNative As Boolean = False)
    Dim encodedMetadata As String
    Dim formulaReference As String
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim sessionId As String
    Dim requestJson As String

    If selectedShape Is Nothing Then
        Err.Raise vbObjectError + 7400, "VisualTeX", "Select exactly one VisualTeX inline formula."
    End If
    encodedMetadata = selectedShape.AlternativeText
    formulaReference = selectedShape.Title
    VTValidateEditEnvelope encodedMetadata, formulaReference, formulaId, displayMode, numbered
    If Len(displayMode) = 0 Then displayMode = "inline"

    VTSetWordMetadataPayload ActiveDocument, formulaId, encodedMetadata
    VTSetWordFormulaFormat ActiveDocument, formulaId, displayMode, numbered

    sessionId = VTNewUuidV4()
    requestJson = VTRequestJson( _
        sessionId, _
        VT_WORD_HOST, _
        "edit", _
        formulaId, _
        displayMode, _
        numbered, _
        VTWordDocumentIdentity(), _
        encodedMetadata, _
        encodedMetadata, _
        "", _
        "", _
        convertToNative)
    VTWriteRequest sessionId, requestJson
    VTLaunchSession VT_WORD_HOST, sessionId
End Sub

Private Sub VTWordEditNativeBookmark(ByVal nativeBookmark As Bookmark)
    VTWordOpenNativeSession nativeBookmark
End Sub

Private Sub VTWordOpenNativeSession(ByVal nativeBookmark As Bookmark)

    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim sessionId As String
    Dim requestJson As String
    Dim nativeMath As OMath

    If nativeBookmark Is Nothing Then
        Err.Raise vbObjectError + 7452, "VisualTeX", "Select one VisualTeX native Word equation."
    End If
    If Not VTTryFormulaIdFromNativeBookmark(nativeBookmark.Name, formulaId) Then
        Err.Raise vbObjectError + 7453, "VisualTeX", "The selected VisualTeX native equation bookmark is invalid."
    End If
    Set nativeMath = VTNativeMathForBookmark(nativeBookmark)
    If nativeMath Is Nothing Then
        Err.Raise vbObjectError + 7454, "VisualTeX", "The selected VisualTeX bookmark no longer contains exactly one native equation."
    End If
    If Not VTTryReadWordMetadataPayload(ActiveDocument, formulaId, encodedMetadata) Then
        Err.Raise vbObjectError + 7455, "VisualTeX", "The selected native equation is missing VisualTeX edit metadata."
    End If
    If Not VTTryReadWordFormulaFormat( _
        ActiveDocument, formulaId, displayMode, numbered) Then
        Err.Raise vbObjectError + 7456, "VisualTeX", "The selected native equation is missing its VisualTeX display format."
    End If

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
        True)
    VTWriteRequest sessionId, requestJson
    VTLaunchSession VT_WORD_HOST, sessionId
End Sub

Public Sub VisualTeX_ConvertSelectedToNativeEquation()
    On Error GoTo Failed

    VTRequireWritableWordDocument
    If Selection.InlineShapes.Count <> 1 Or _
       Not VTIsVisualTeXInlineShape(Selection.InlineShapes(1)) Then
        Err.Raise vbObjectError + 7428, "VisualTeX", _
            "Select exactly one VisualTeX formula image to convert to Word OMML."
    End If
    VTWordConvertInlineShapeToNativeEquation Selection.InlineShapes(1)
    Exit Sub

Failed:
    VTShowError "Word native equation conversion", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_UpdateEquationNumbers()
    Dim updated As Long

    On Error GoTo Failed
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7401, "VisualTeX", "Open a Word document first."
    End If
    updated = VTPruneOrphanedEquationNumberScaffolds(ActiveDocument)
    VTReconcileEquationNumbers ActiveDocument
    updated = VTVariantArrayCount( _
        VTValidNumberedFormulaIds(ActiveDocument))
    VTShowInformation "Updated " & CStr(updated) & _
        " VisualTeX equation numbers."
    Exit Sub

Failed:
    VTShowError "equation numbering", Err.Number, Err.Description
End Sub

Private Function VTEquationSequenceFieldByIndex( _
    ByVal documentObject As Document, _
    ByVal itemIndex As Long) As Field

    Dim candidate As Field
    Dim equationLabelName As String
    Dim bookmarkName As String
    Dim currentIndex As Long

    If documentObject Is Nothing Or itemIndex <= 0 Then Exit Function
    equationLabelName = VTNativeEquationLabelName()
    For Each candidate In documentObject.Fields
        If VTIsNativeEquationSequenceField(candidate, equationLabelName) Then
            bookmarkName = VTEquationNumberBookmarkForField( _
                documentObject, candidate)
            If Len(bookmarkName) > 0 Then
                currentIndex = currentIndex + 1
                If currentIndex = itemIndex Then
                    Set VTEquationSequenceFieldByIndex = candidate
                    Exit Function
                End If
            End If
        End If
    Next candidate
End Function

Private Function VTEquationNumberBookmarkForField( _
    ByVal documentObject As Document, _
    ByVal sequenceField As Field) As String

    Dim candidate As Bookmark
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim matchCount As Long
    Dim suffixText As String
    Dim visibleBookmarkName As String

    If documentObject Is Nothing Or sequenceField Is Nothing Then Exit Function
    For Each candidate In documentObject.Bookmarks
        If Left$(candidate.Name, _
           Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX Then
            If candidate.Range.Start <= sequenceField.Result.Start And _
               candidate.Range.End >= sequenceField.Result.End Then
                suffixText = Mid$(candidate.Name, _
                    Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX) + 1)
                visibleBookmarkName = _
                    VT_WORD_NUMBER_BOOKMARK_PREFIX & suffixText
                If documentObject.Bookmarks.Exists(visibleBookmarkName) Then
                    VTEquationNumberBookmarkForField = visibleBookmarkName
                    Exit Function
                End If
            End If
        End If
    Next candidate
    fieldStart = VTEquationFieldStart(sequenceField) - 1
    fieldEnd = VTEquationFieldEnd(sequenceField) + 1
    For Each candidate In documentObject.Bookmarks
        If Left$(candidate.Name, Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_NUMBER_BOOKMARK_PREFIX Then
            If candidate.Range.Start <= fieldStart And _
               candidate.Range.End >= fieldEnd Then
                matchCount = matchCount + 1
                VTEquationNumberBookmarkForField = candidate.Name
            End If
        End If
    Next candidate
    If matchCount <> 1 Then VTEquationNumberBookmarkForField = ""
End Function

Private Function VTEquationCrossReferenceLabel( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal ordinal As Long) As String

    Dim latexBase64 As String
    Dim previewText As String

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Or _
       ordinal < 1 Then Exit Function
    previewText = "VisualTeX formula"
    If VTTryReadWordLatexPayload( _
       documentObject, formulaId, latexBase64) Then
        previewText = VTEquationCrossReferenceText(latexBase64)
    End If
    VTEquationCrossReferenceLabel = _
        "(" & CStr(ordinal) & ")  " & previewText
End Function

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

Private Function VTEquationNumberCrossReferenceItems( _
    ByVal documentObject As Document) As Variant

    Dim formulaIds As Variant
    Dim itemIndex As Long
    Dim itemCount As Long
    Dim sequenceBookmarkName As String
    Dim numberText As String
    Dim items() As String

    If documentObject Is Nothing Then Exit Function
    VTPruneOrphanedEquationNumberScaffolds documentObject
    VTReconcileEquationNumbers documentObject
    formulaIds = VTValidNumberedFormulaIds(documentObject)
    itemCount = VTVariantArrayCount(formulaIds)
    If itemCount <= 0 Then Exit Function

    ReDim items(1 To itemCount)
    For itemIndex = 1 To itemCount
        sequenceBookmarkName = VTEquationSequenceNumberBookmarkName( _
            CStr(formulaIds(itemIndex)))
        If Not documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
            Err.Raise vbObjectError + 7547, "VisualTeX", _
                "A VisualTeX formula has no live Equation number target."
        End If
        numberText = Trim$(documentObject.Bookmarks( _
            sequenceBookmarkName).Range.Text)
        items(itemIndex) = VTEquationCrossReferenceLabel( _
            documentObject, CStr(formulaIds(itemIndex)), CLng(numberText))
    Next itemIndex
    VTEquationNumberCrossReferenceItems = items
End Function

Private Function VTInsertEquationNumberReferenceAtRange( _
    ByVal targetRange As Range, _
    ByVal itemIndex As Long) As Range

    Dim documentObject As Document
    Dim formulaIds As Variant
    Dim formulaId As String
    Dim sequenceBookmarkName As String
    Dim expectedNumber As String
    Dim referenceField As Field
    Dim insertedRange As Range
    Dim fieldRange As Range
    Dim insertionStart As Long
    Dim insertionEnd As Long
    Dim itemCount As Long

    If targetRange Is Nothing Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The Equation cross-reference insertion Range is missing."
    End If
    Set documentObject = targetRange.Document
    VTPruneOrphanedEquationNumberScaffolds documentObject
    VTReconcileEquationNumbers documentObject
    formulaIds = VTValidNumberedFormulaIds(documentObject)
    itemCount = VTVariantArrayCount(formulaIds)
    If itemIndex < 1 Or itemIndex > itemCount Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The selected VisualTeX Equation item does not exist."
    End If

    formulaId = CStr(formulaIds(itemIndex))
    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The selected VisualTeX Equation has no live number target."
    End If
    expectedNumber = Trim$( _
        documentObject.Bookmarks(sequenceBookmarkName).Range.Text)

    ' Insert a native REF directly to the exact VT_N_ SEQ result Bookmark.
    ' Word's built-in caption cross-reference targets an entire single-paragraph
    ' equation line, which would copy tabs and the formula itself. A direct REF
    ' stays fully dynamic while returning only the live number.
    insertionStart = targetRange.Start
    Set insertedRange = documentObject.Range( _
        Start:=insertionStart, End:=insertionStart)
    insertedRange.Text = "()"
    Set fieldRange = documentObject.Range( _
        Start:=insertionStart + 1, End:=insertionStart + 1)
    Set referenceField = documentObject.Fields.Add( _
        Range:=fieldRange, Type:=wdFieldRef, _
        Text:=VTParenthesizedEquationReferenceFieldText( _
            sequenceBookmarkName), _
        PreserveFormatting:=False)
    referenceField.Update
    insertionEnd = VTEquationFieldEnd(referenceField) + 1
    Set insertedRange = documentObject.Range( _
        Start:=insertionStart, End:=insertionEnd)
    VTFormatBodyEquationReference _
        documentObject, referenceField, insertedRange
    If VTFirstPositiveIntegerInText(referenceField.Result.Text) <> _
       CLng(expectedNumber) Or _
       insertedRange.Text <> "(" & expectedNumber & ")" Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The inserted native Equation cross-reference is incomplete" & _
            " [code=" & referenceField.Code.Text & _
            "; result=" & referenceField.Result.Text & _
            "; text=" & insertedRange.Text & "]."
    End If
    Set VTInsertEquationNumberReferenceAtRange = insertedRange.Duplicate
End Function

Public Sub VisualTeX_OpenEquationCrossReference()
    Dim crossReferenceItems As Variant
    Dim selectedIndex As Variant
    Dim promptText As String
    Dim itemText As String
    Dim itemIndex As Long
    Dim itemCount As Long
    Dim displayCount As Long
    Dim insertedRange As Range

    On Error GoTo Failed
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7401, "VisualTeX", "Open a Word document first."
    End If
    crossReferenceItems = _
        VTEquationNumberCrossReferenceItems(ActiveDocument)
    If Not IsArray(crossReferenceItems) Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "This document has no Equation cross-reference items."
    End If
    itemCount = UBound(crossReferenceItems) - _
        LBound(crossReferenceItems) + 1
    If itemCount <= 0 Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "This document has no Equation cross-reference items."
    End If

    displayCount = itemCount
    If displayCount > 15 Then displayCount = 15
    promptText = "Enter the VisualTeX equation list number:" & _
        vbCrLf & vbCrLf
    For itemIndex = 1 To displayCount
        itemText = CStr(crossReferenceItems( _
            LBound(crossReferenceItems) + itemIndex - 1))
        itemText = Replace$(Replace$(itemText, vbTab, " "), vbCr, " ")
        If Len(itemText) > 72 Then itemText = Left$(itemText, 69) & "..."
        promptText = promptText & CStr(itemIndex) & ". " & itemText & vbCrLf
    Next itemIndex
    If itemCount > displayCount Then
        promptText = promptText & "... total " & CStr(itemCount) & _
            " VisualTeX equations" & vbCrLf
    End If

    selectedIndex = InputBox( _
        Prompt:=promptText, _
        Title:="VisualTeX Equation Reference")
    If Len(Trim$(CStr(selectedIndex))) = 0 Then Exit Sub
    If Not IsNumeric(selectedIndex) Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "Enter the numeric list index of a VisualTeX equation."
    End If
    itemIndex = CLng(selectedIndex)
    If itemIndex < 1 Or itemIndex > itemCount Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "Enter a VisualTeX equation list index between 1 and " & _
            CStr(itemCount) & "."
    End If

    Set insertedRange = VTInsertEquationNumberReferenceAtRange( _
        Selection.Range.Duplicate, itemIndex)
    insertedRange.Collapse wdCollapseEnd
    insertedRange.Select
    Exit Sub

Failed:
    VTShowError "equation cross-reference", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_OpenApplication()
    On Error GoTo Failed
    VTOpenApplication VT_WORD_HOST
    Exit Sub
Failed:
    VTShowError "application launch", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_ApplyPendingResult()
    Dim sessionId As String
    Dim dispatch As Object
    Dim actionName As String
    Dim hostName As String

    On Error GoTo Failed
    sessionId = VTReadActiveSessionId(VT_WORD_HOST)
    VTTraceWordSession sessionId, "callback-enter", ""
    Set dispatch = VTReadDispatch(sessionId)
    actionName = CStr(dispatch("action"))
    hostName = CStr(dispatch("host"))
    If hostName <> VT_WORD_HOST Then
        Err.Raise vbObjectError + 7402, "VisualTeX", "The active VisualTeX dispatch is not for Word."
    End If

    Select Case actionName
        Case "commit": VTCommitWordDispatch sessionId, dispatch
        Case "cancel": VTCancelWordDispatch sessionId, dispatch
        Case Else
            Err.Raise vbObjectError + 7403, "VisualTeX", "The VisualTeX Word dispatch action is invalid."
    End Select
    Exit Sub

Failed:
    Err.Raise Err.Number, "VisualTeX Word callback", Err.Description
End Sub

Private Function VTPrepareWordCreateInsertionRange( _
    ByVal requestedRange As Range, _
    ByVal displayMode As String) As Range

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

    If requestedRange Is Nothing Then
        Err.Raise vbObjectError + 7551, "VisualTeX", _
            "The Word formula insertion Range is missing."
    End If
    If displayMode <> "inline" And displayMode <> "block" Then
        Err.Raise vbObjectError + 7451, "VisualTeX", _
            "The VisualTeX Word display mode is invalid."
    End If

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
    If insertionRange.Information(wdWithInTable) Then
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
    Else
        VTNormalizePlainWordParagraph targetParagraph
    End If
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

Private Sub VTNormalizePlainWordParagraph(ByVal paragraphRange As Range)
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

Private Function VTInsertDedicatedEquationHelperParagraph( _
    ByVal layoutTable As Table) As Range

    Dim documentObject As Document
    Dim tableRange As Range
    Dim helperParagraph As Range
    Dim helperStart As Long

    If layoutTable Is Nothing Then
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "The Equation helper paragraph requires a numbered table."
    End If
    Set documentObject = layoutTable.Range.Document
    helperStart = layoutTable.Range.End
    Set tableRange = layoutTable.Range.Duplicate

    ' Always create a new paragraph after the entire numbered table. A collapsed
    ' Range at Table.Range.End can resolve into the existing following body line
    ' on Word for Mac, which previously exposed an internal helper marker and
    ' rejected a visually empty insertion position. InsertParagraphAfter preserves that body
    ' line and gives the native SEQ field its own paragraph unconditionally.
    tableRange.InsertParagraphAfter
    Set helperParagraph = documentObject.Range( _
        Start:=helperStart, End:=helperStart).Paragraphs(1).Range.Duplicate
    If helperParagraph.Start <> helperStart Or _
       helperParagraph.Information(wdWithInTable) Or _
       helperParagraph.Fields.Count <> 0 Or _
       helperParagraph.InlineShapes.Count <> 0 Or _
       helperParagraph.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(helperParagraph) Then
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "Word did not create an independent empty Equation helper paragraph."
    End If
    Set VTInsertDedicatedEquationHelperParagraph = helperParagraph
End Function

Private Function VTEnsurePlainContinuationParagraph( _
    ByVal sourceParagraph As Range) As Range

    Dim documentObject As Document
    Dim insertionRange As Range
    Dim probeRange As Range
    Dim markerRange As Range
    Dim continuationParagraph As Range
    Dim continuationStart As Long
    Dim markerStart As Long

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
    ElseIf continuationParagraph.Information(wdWithInTable) Or _
           continuationParagraph.Fields.Count > 0 Or _
           continuationParagraph.InlineShapes.Count > 0 Or _
           continuationParagraph.OMaths.Count > 0 Then
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
    VTNormalizePlainWordParagraph continuationParagraph
    Set VTEnsurePlainContinuationParagraph = continuationParagraph
End Function

Private Sub VTPlaceCaretAfterDisplayFormula( _
    ByVal formulaRange As Range, _
    ByVal formulaId As String)

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
    caretRange.Font.Position = 0
    caretRange.Font.Hidden = False
    caretRange.Font.Color = wdColorAutomatic
    caretRange.Select
End Sub

Private Sub VTWordCreate( _
    ByVal displayMode As String, _
    ByVal numbered As Boolean, _
    Optional ByVal nativeEquation As Boolean = False)
    Dim sessionId As String
    Dim formulaId As String
    Dim pendingMarker As String
    Dim placeholder As InlineShape
    Dim insertionRange As Range
    Dim requestJson As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed
    VTRequireWritableWordDocument
    VTRefreshWordHealthQuietly
    VTCleanupOrphanedNumberedDisplaySelection Selection.Range
    If Not VTPathFileExists(VTPlaceholderImagePath()) Then
        Err.Raise vbObjectError + 7404, "VisualTeX", "The VisualTeX placeholder resource is missing. Repair the offline add-in."
    End If

    sessionId = VTNewUuidV4()
    formulaId = VTNewUuidV4()
    pendingMarker = VTPendingMarker(sessionId, formulaId)
    Set insertionRange = VTPrepareWordCreateInsertionRange( _
        Selection.Range.Duplicate, displayMode)
    Set placeholder = ActiveDocument.InlineShapes.AddPicture( _
        FileName:=VTPlaceholderImagePath(), _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Range:=insertionRange)
    placeholder.AlternativeText = pendingMarker
    placeholder.Title = pendingMarker
    placeholder.Width = 1
    placeholder.Height = 1
    VTAddPendingBookmark placeholder.Range, sessionId
    VTTraceWordSession sessionId, "placeholder-created", pendingMarker
    Selection.SetRange Start:=placeholder.Range.End, End:=placeholder.Range.End

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
        nativeEquation)
    VTWriteRequest sessionId, requestJson
    VTTraceWordSession sessionId, "request-written", pendingMarker
    VTLaunchSession VT_WORD_HOST, sessionId
    VTTraceWordSession sessionId, "editor-launched", pendingMarker
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Not placeholder Is Nothing Then placeholder.Delete
    If Len(sessionId) > 0 Then VTDeleteSessionFiles sessionId
    On Error GoTo 0
    VTShowError "Word formula creation", errorNumber, errorDescription
End Sub

Private Sub VTCommitWordDispatch(ByVal sessionId As String, ByVal dispatch As Object)
    Dim mode As String
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim nativeEquation As Boolean
    Dim nativeDisplayMode As String
    Dim imagePath As String
    Dim metadata As String
    Dim latexBase64 As String
    Dim ommlBase64 As String
    Dim nativeDocumentPath As String
    Dim pendingMarker As String
    Dim sourceMarker As String
    Dim sourceDocumentId As String
    Dim targetDocument As Document
    Dim targetImage As InlineShape
    Dim pendingBookmark As Bookmark
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
    Dim nativeTargetReplaced As Boolean
    Dim nativeBookmarkSet As Boolean
    Dim targetIsNative As Boolean
    Dim committed As InlineShape
    Dim candidate As InlineShape
    Dim insertionRange As Range
    Dim widthPoints As Double
    Dim heightPoints As Double
    Dim baselinePoints As Double
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
    Dim hadPreviousLatexPayload As Boolean
    Dim hadPreviousOmmlPayload As Boolean
    Dim hadPreviousMetadataPayload As Boolean
    Dim hadPreviousFormat As Boolean
    Dim formulaStateStored As Boolean
    Dim deferNativeDisplay As Boolean
    Dim pendingPlaceholderRemoved As Boolean
    Dim pendingPlaceholderStart As Long
    Dim nativeEquationStart As Long
    Dim restoredPlaceholder As InlineShape
    Dim targetFormulaId As String
    Dim transactionErrorNumber As Long
    Dim transactionErrorDescription As String
    Dim transactionStage As String
    Dim internalMutationStarted As Boolean

    transactionStage = "validate-document"
    VTRequireWritableWordDocument
    Set targetDocument = ActiveDocument
    VTRequireDispatchValue dispatch, "mode"
    VTRequireDispatchValue dispatch, "formulaId"
    VTRequireDispatchValue dispatch, "displayMode"
    VTRequireDispatchValue dispatch, "numbered"
    VTRequireDispatchValue dispatch, "imagePath"
    VTRequireDispatchValue dispatch, "metadata"
    VTRequireDispatchValue dispatch, "latexBase64"
    VTRequireDispatchValue dispatch, "ommlBase64"

    transactionStage = "read-dispatch"
    mode = CStr(dispatch("mode"))
    formulaId = CStr(dispatch("formulaId"))
    displayMode = CStr(dispatch("displayMode"))
    numbered = (CStr(dispatch("numbered")) = "1")
    nativeEquation = (VTDispatchOptional(dispatch, "nativeEquation") = "1")
    imagePath = CStr(dispatch("imagePath"))
    metadata = CStr(dispatch("metadata"))
    latexBase64 = CStr(dispatch("latexBase64"))
    ommlBase64 = CStr(dispatch("ommlBase64"))
    nativeDocumentPath = VTDispatchOptional(dispatch, "nativeDocumentPath")
    pendingMarker = VTDispatchOptional(dispatch, "pendingMarker")
    sourceMarker = VTDispatchOptional(dispatch, "sourceMarker")
    sourceDocumentId = VTDispatchOptional(dispatch, "sourceDocumentId")

    VTTraceWordSession sessionId, "commit-dispatch-read", pendingMarker

    If Len(sourceDocumentId) = 0 Or sourceDocumentId <> VTWordDocumentIdentity() Then
        Err.Raise vbObjectError + 7415, "VisualTeX", "The active Word document changed while VisualTeX was open."
    End If
    widthPoints = VTDispatchPositiveDouble(dispatch, "widthPoints")
    heightPoints = VTDispatchPositiveDouble(dispatch, "heightPoints")
    baselinePoints = VTDispatchOptionalDouble(dispatch, "baseline", 0#)

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

    formulaReference = VTFormulaReference(formulaId, displayMode, numbered)
    captionText = VTEquationCrossReferenceText(latexBase64)
    VTValidateAbsoluteVisualTeXPath imagePath
    If Not VTPathFileExists(imagePath) Then
        Err.Raise vbObjectError + 7406, "VisualTeX", "VisualTeX Word result image is missing."
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
    Else
        Set targetImage = VTFindUniqueInlineShape(sourceMarker)
    End If
    Err.Clear
    On Error GoTo RollbackCandidate
    VTBeginWordInternalMutation
    internalMutationStarted = True
    VTTraceWordSession sessionId, "commit-target-resolved", pendingMarker

    transactionStage = "capture-target-range"
    If targetIsNative Then
        Set originalNativeMath = VTNativeMathForBookmark(nativeTarget)
        If Not VTTryFormulaIdFromNativeBookmark(nativeTarget.Name, targetFormulaId) Or _
           targetFormulaId <> formulaId Or originalNativeMath Is Nothing Then
            Err.Raise vbObjectError + 7454, "VisualTeX", "The original VisualTeX native equation target is invalid."
        End If
        Set originalNativeRange = originalNativeMath.Range.Duplicate
        originalNativeStart = originalNativeRange.Start
        originalNativeBookmarkName = nativeTarget.Name
        Set targetRange = originalNativeRange.Duplicate
    ElseIf Not targetImage Is Nothing Then
        Set targetRange = targetImage.Range.Duplicate
    ElseIf targetFromPendingBookmark And Not pendingBookmark Is Nothing Then
        Set targetRange = pendingBookmark.Range.Duplicate
    Else
        If nativeEquation Then
            If targetDocument.Bookmarks.Exists(VTNativeFormulaBookmarkName(formulaId)) Then
                VTSetWordLatexPayload targetDocument, formulaId, latexBase64
                VTSetWordOmmlPayload targetDocument, formulaId, ommlBase64
                VTSetWordMetadataPayload targetDocument, formulaId, metadata
                VTSetWordFormulaFormat targetDocument, formulaId, displayMode, numbered
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
                VTDeletePendingBookmark targetDocument, sessionId
                GoTo CommitSucceeded
            End If
        End If
        Err.Raise vbObjectError + 7426, "VisualTeX", "The original Word formula is missing and no committed VisualTeX result was found."
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

    If nativeEquation Then
        transactionStage = "prepare-native-replacement"
        nativeDisplayMode = displayMode
        ' Word for Mac is significantly more stable when an unnumbered display
        ' equation is first transferred as inline OMath and promoted only after
        ' the source image/placeholder transaction has settled. Apply this to
        ' creates and edits alike so a stale display Range cannot absorb nearby
        ' content or leave a broken editor Session behind.
        If displayMode = "block" And Not numbered Then
            deferNativeDisplay = True
            nativeDisplayMode = "inline"
        ElseIf numbered Then
            nativeDisplayMode = "inline"
        End If

        ' Replacing a native equation by first inserting beside it is unsafe in
        ' Word for Mac: Word can merge both equations and mutate the original
        ' COM Range so a later Delete consumes adjacent body text. Back up the
        ' original OMath in a hidden document and replace its exact Range in
        ' one FormattedText assignment instead.
        If targetIsNative Then
            Set originalNativeBackupDocument = Documents.Add(Visible:=False)
            Set originalNativeBackupRange = originalNativeBackupDocument.Content
            originalNativeBackupRange.Collapse wdCollapseStart
            originalNativeBackupRange.FormattedText = originalNativeRange.FormattedText
            If originalNativeBackupDocument.OMaths.Count <> 1 Then
                Err.Raise vbObjectError + 7471, "VisualTeX", _
                    "Word could not back up the original native equation before replacement."
            End If
            Set originalNativeBackupRange = _
                originalNativeBackupDocument.OMaths(1).Range.Duplicate
            targetRange.Document.Activate
        End If

        transactionStage = "insert-native-equation"
        Set nativeEquationRange = VTInsertNativeEquationAtRange( _
            targetRange, _
            ommlBase64, _
            nativeDocumentPath, _
            nativeDisplayMode, _
            displayMode = "block", _
            targetIsNative)
        nativeEquationStart = nativeEquationRange.Start
        nativeTargetReplaced = targetIsNative

        transactionStage = "store-native-state"
        targetDocument.Activate
        VTSetWordLatexPayload targetDocument, formulaId, latexBase64
        VTSetWordOmmlPayload targetDocument, formulaId, ommlBase64
        VTSetWordMetadataPayload targetDocument, formulaId, metadata
        VTSetWordFormulaFormat targetDocument, formulaId, displayMode, numbered
        formulaStateStored = True

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
            targetImage.Delete
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
        If originalNativeMath Is Nothing Then
            Err.Raise vbObjectError + 7460, "VisualTeX", _
                "Word could not resolve the newly inserted native equation identity."
        End If
        Set nativeEquationRange = originalNativeMath.Range.Duplicate

        If displayMode = "block" Then
            If numbered Then
                transactionStage = "normalize-native-number-layout"
                numberCreated = False
                Set numberLayoutRange = VTEnsureNativeEquationNumber( _
                    nativeEquationRange, heightPoints, formulaId, captionText, _
                    numberCreated)
                If numberCreated Then Set insertedNumber = numberLayoutRange
                Set nativeEquationRange = VTResolveNativeEquationRange( _
                    targetDocument, nativeEquationStart, 16)
            Else
                transactionStage = "promote-native-display"
                Set nativeEquationRange = _
                    VTPromoteNativeEquationToDisplay(nativeEquationRange)
            End If
        Else
            transactionStage = "finalize-native-inline"
            Set nativeEquationRange = _
                VTFinalizeInlineNativeEquation(nativeEquationRange)
        End If

        transactionStage = "bookmark-native-equation"
        VTSetNativeFormulaBookmark targetDocument, nativeEquationRange, formulaId
        nativeBookmarkSet = True
        VTDeletePendingBookmark targetDocument, sessionId

        On Error Resume Next
        If Not originalNativeBackupDocument Is Nothing Then
            originalNativeBackupDocument.Close SaveChanges:=wdDoNotSaveChanges
        End If
        On Error GoTo RollbackCandidate
        transactionStage = "place-native-caret"
        On Error Resume Next
        If displayMode = "inline" Then
            VTPlaceCaretAfterInlineNativeEquation nativeEquationRange
        ElseIf mode = "create" Then
            VTPlaceCaretAfterDisplayFormula nativeEquationRange, formulaId
        Else
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
        GoTo CommitSucceeded
    End If

    If targetIsNative Then
        Set originalNativeBackupDocument = Documents.Add(Visible:=False)
        Set originalNativeBackupRange = originalNativeBackupDocument.Content
        originalNativeBackupRange.Collapse wdCollapseStart
        originalNativeBackupRange.FormattedText = originalNativeRange.FormattedText
        If originalNativeBackupDocument.OMaths.Count <> 1 Then
            Err.Raise vbObjectError + 7471, "VisualTeX", _
                "Word could not back up the original native equation before image replacement."
        End If
        Set originalNativeBackupRange = _
            originalNativeBackupDocument.OMaths(1).Range.Duplicate
        targetRange.Document.Activate
    End If

    Set insertionRange = targetRange.Duplicate
    If Not targetIsNative Then insertionRange.Collapse wdCollapseStart
    targetDocument.Activate
    Set candidate = targetDocument.InlineShapes.AddPicture( _
        FileName:=imagePath, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Range:=insertionRange)
    nativeTargetReplaced = targetIsNative
    candidate.LockAspectRatio = msoFalse
    candidate.Width = CSng(widthPoints)
    candidate.Height = CSng(heightPoints)
    candidate.LockAspectRatio = msoTrue
    candidate.AlternativeText = metadata
    candidate.Title = formulaReference
    If displayMode = "inline" Then
        If baselinePoints > 0# Or baselinePoints < -256# Then
            Err.Raise vbObjectError + 7408, "VisualTeX", "VisualTeX Word baseline is outside the allowed range."
        End If
        candidate.Range.Font.Position = CLng(baselinePoints)
    ElseIf Not (targetIsNative And numbered) Then
        candidate.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If

    If Abs(candidate.Width - widthPoints) > 0.1 Or Abs(candidate.Height - heightPoints) > 0.1 Or _
       candidate.AlternativeText <> metadata Or candidate.Title <> formulaReference Then
        Err.Raise vbObjectError + 7422, "VisualTeX", "Word did not persist the VisualTeX formula properties."
    End If
    If displayMode = "inline" Then
        If candidate.Range.Font.Position <> CLng(baselinePoints) Then
            Err.Raise vbObjectError + 7423, "VisualTeX", "Word did not persist the VisualTeX inline baseline."
        End If
    ElseIf Not (targetIsNative And numbered) Then
        If candidate.Range.ParagraphFormat.Alignment <> wdAlignParagraphCenter Then
            Err.Raise vbObjectError + 7424, "VisualTeX", "Word did not persist the VisualTeX display alignment."
        End If
    End If

    VTSetWordLatexPayload targetDocument, formulaId, latexBase64
    VTSetWordOmmlPayload targetDocument, formulaId, ommlBase64
    VTSetWordMetadataPayload targetDocument, formulaId, metadata
    VTSetWordFormulaFormat targetDocument, formulaId, displayMode, numbered
    formulaStateStored = True

    ' Finalize the paragraph only after the old image/native target has gone.
    ' Otherwise the placeholder participates in the tabbed line and Word shifts
    ' the formula away from the true text-column center when it is deleted.
    transactionStage = "remove-image-source"
    If targetIsNative Then
        If targetRange.Document.Bookmarks.Exists(originalNativeBookmarkName) Then
            targetRange.Document.Bookmarks(originalNativeBookmarkName).Delete
        End If
        originalNativeBookmarkDeleted = True
    ElseIf Not targetImage Is Nothing Then
        targetImage.Delete
    End If

    transactionStage = "resolve-image-after-source-removal"
    Set candidate = VTFindCommittedInlineShape(metadata, formulaReference)
    If candidate Is Nothing Then
        Err.Raise vbObjectError + 7426, "VisualTeX", _
            "Word could not resolve the committed formula image after replacement."
    End If

    If displayMode = "block" Then
        If numbered Then
            transactionStage = "normalize-image-number-layout"
            numberCreated = False
            Set numberLayoutRange = VTEnsureImageEquationNumber( _
                candidate, heightPoints, formulaId, captionText, numberCreated)
            If numberCreated Then Set insertedNumber = numberLayoutRange
        Else
            transactionStage = "normalize-image-display-layout"
            VTNormalizeUnnumberedDisplayParagraph candidate.Range
        End If
    End If
    VTDeletePendingBookmark targetDocument, sessionId

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
    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    Exit Sub

RollbackCandidate:
    transactionErrorNumber = Err.Number
    transactionErrorDescription = Err.Description
    VTWriteWordFailureTrace _
        sessionId, transactionStage, transactionErrorNumber, transactionErrorDescription
    On Error Resume Next
    If Not insertedNumber Is Nothing Then insertedNumber.Delete
    If nativeBookmarkSet Then
        If targetRange.Document.Bookmarks.Exists( _
            VTNativeFormulaBookmarkName(formulaId)) Then
            targetRange.Document.Bookmarks( _
                VTNativeFormulaBookmarkName(formulaId)).Delete
        End If
    End If

    If targetIsNative And nativeTargetReplaced And _
       Not originalNativeBackupRange Is Nothing Then
        If Not candidate Is Nothing Then
            Set rollbackRange = candidate.Range.Duplicate
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
            End If
        End If
    Else
        If Not candidate Is Nothing Then candidate.Delete
        If Not nativeEquationRange Is Nothing Then nativeEquationRange.Delete
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
    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    On Error GoTo 0
    Err.Raise transactionErrorNumber, "VisualTeX Word transaction", _
        transactionStage & ": " & transactionErrorDescription
End Sub

Private Sub VTCancelWordDispatch(ByVal sessionId As String, ByVal dispatch As Object)
    Dim pendingMarker As String
    Dim sourceDocumentId As String
    Dim pendingBookmark As Bookmark
    Dim target As InlineShape

    pendingMarker = VTDispatchOptional(dispatch, "pendingMarker")
    sourceDocumentId = VTDispatchOptional(dispatch, "sourceDocumentId")
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
End Sub

Private Function VTFindUniqueInlineShape(ByVal marker As String) As InlineShape
    Dim item As InlineShape
    Dim match As InlineShape
    Dim count As Long

    If Len(marker) = 0 Or Len(marker) > 131072 Then
        Err.Raise vbObjectError + 7409, "VisualTeX", "VisualTeX Word target marker is invalid."
    End If
    For Each item In ActiveDocument.InlineShapes
        If item.AlternativeText = marker Or item.Title = marker Then
            count = count + 1
            Set match = item
        End If
    Next item
    If count <> 1 Then
        Err.Raise vbObjectError + 7410, "VisualTeX", "Word must contain exactly one matching VisualTeX formula object."
    End If
    Set VTFindUniqueInlineShape = match
End Function

Private Function VTFindCommittedInlineShape(ByVal metadata As String, ByVal formulaReference As String) As InlineShape
    Dim item As InlineShape
    Dim match As InlineShape
    Dim count As Long

    For Each item In ActiveDocument.InlineShapes
        If item.AlternativeText = metadata And item.Title = formulaReference Then
            count = count + 1
            Set match = item
        End If
    Next item
    If count > 1 Then
        Err.Raise vbObjectError + 7427, "VisualTeX", "Word contains multiple committed copies of the same VisualTeX Session result."
    End If
    If count = 1 Then Set VTFindCommittedInlineShape = match
End Function

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
    Dim formulaStart As Long
    Dim sourceAlternativeText As String
    Dim sourceTitle As String
    Dim operationErrorNumber As Long
    Dim operationErrorDescription As String

    If formulaRange Is Nothing Or formulaRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7542, "VisualTeX", _
            "The numbered formula image backup target is invalid."
    End If
    Set documentObject = formulaRange.Document
    formulaStart = formulaRange.Start
    sourceAlternativeText = formulaRange.InlineShapes(1).AlternativeText
    sourceTitle = formulaRange.InlineShapes(1).Title
    On Error GoTo RestoreFailed

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
        Start:=formulaStart + 1, End:=formulaStart + 2)
    If restoredRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7542, "VisualTeX", _
            "Word could not restore the numbered formula image after the center tab."
    End If
    restoredRange.InlineShapes(1).AlternativeText = sourceAlternativeText
    restoredRange.InlineShapes(1).Title = sourceTitle
    If restoredRange.InlineShapes(1).AlternativeText <> _
       sourceAlternativeText Or _
       restoredRange.InlineShapes(1).Title <> sourceTitle Then
        Err.Raise vbObjectError + 7542, "VisualTeX", _
            "Word did not preserve the VisualTeX image identity after layout."
    End If
    Set VTPrependCenterTabPreservingImage = _
        restoredRange.InlineShapes(1).Range.Duplicate
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

Private Function VTPrependCenterTabPreservingNativeFormula( _
    ByVal formulaRange As Range, _
    ByVal formulaId As String) As Range

    Dim documentObject As Document
    Dim backupDocument As Document
    Dim backupRange As Range
    Dim insertionRange As Range
    Dim restoredRange As Range
    Dim rollbackMath As OMath
    Dim cleanupRange As Range
    Dim bookmarkName As String
    Dim formulaStart As Long
    Dim sourceDeleted As Boolean
    Dim operationStage As String
    Dim operationErrorNumber As Long
    Dim operationErrorDescription As String

    If formulaRange Is Nothing Or formulaRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7545, "VisualTeX", _
            "The numbered native formula backup target is invalid."
    End If
    Set formulaRange = formulaRange.OMaths(1).Range.Duplicate
    Set documentObject = formulaRange.Document
    formulaStart = formulaRange.Start
    bookmarkName = VTNativeFormulaBookmarkName(formulaId)
    On Error GoTo RestoreFailed

    operationStage = "backup-native-formula"
    Set backupDocument = Documents.Add(Visible:=False)
    Set backupRange = backupDocument.Content
    backupRange.Collapse wdCollapseStart
    backupRange.FormattedText = formulaRange.FormattedText
    If backupDocument.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7545, "VisualTeX", _
            "Word could not back up the numbered native formula."
    End If
    Set backupRange = backupDocument.OMaths(1).Range.Duplicate

    operationStage = "remove-native-bookmark"
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        documentObject.Bookmarks(bookmarkName).Delete
    End If

    operationStage = "remove-native-formula"
    documentObject.Activate
    formulaRange.Delete
    sourceDeleted = True

    operationStage = "insert-center-tab"
    Set insertionRange = documentObject.Range( _
        Start:=formulaStart, End:=formulaStart)
    insertionRange.InsertBefore vbTab

    operationStage = "restore-native-formula"
    Set insertionRange = documentObject.Range( _
        Start:=formulaStart + 1, End:=formulaStart + 1)
    insertionRange.FormattedText = backupRange.FormattedText
    Set restoredRange = VTResolveNativeEquationRange( _
        documentObject, formulaStart + 1, 16)
    If restoredRange.Start <> formulaStart + 1 Or _
       restoredRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7545, "VisualTeX", _
            "Word restored the numbered native formula at an invalid boundary."
    End If

    operationStage = "restore-native-bookmark"
    VTSetNativeFormulaBookmark documentObject, restoredRange, formulaId
    Set VTPrependCenterTabPreservingNativeFormula = _
        restoredRange.OMaths(1).Range.Duplicate
    backupDocument.Close SaveChanges:=wdDoNotSaveChanges
    Exit Function

RestoreFailed:
    operationErrorNumber = Err.Number
    operationErrorDescription = Err.Description
    On Error Resume Next
    If sourceDeleted And Not backupRange Is Nothing Then
        Set rollbackMath = VTNativeMathNearStart( _
            documentObject, formulaStart + 1, 16)
        If Not rollbackMath Is Nothing Then rollbackMath.Range.Delete
        If formulaStart < documentObject.Content.End Then
            Set cleanupRange = documentObject.Range( _
                Start:=formulaStart, End:=formulaStart + 1)
            If cleanupRange.Text = vbTab Then cleanupRange.Delete
        End If
        Set insertionRange = documentObject.Range( _
            Start:=formulaStart, End:=formulaStart)
        insertionRange.FormattedText = backupRange.FormattedText
        Set restoredRange = VTResolveNativeEquationRange( _
            documentObject, formulaStart, 16)
        If Not restoredRange Is Nothing Then
            VTSetNativeFormulaBookmark documentObject, restoredRange, formulaId
        End If
    End If
    If Not backupDocument Is Nothing Then
        backupDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise operationErrorNumber, "VisualTeX Equation numbering", _
        "VTPrependCenterTabPreservingNativeFormula/" & operationStage & _
        ": " & operationErrorDescription
End Function

Private Function VTInsertEquationNumber( _
    ByRef formulaShape As InlineShape, _
    ByVal formulaId As String, _
    ByVal captionText As String, _
    Optional ByVal deferReconcile As Boolean = False) As Range

    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim prefixRange As Range
    Dim suffixRange As Range
    Dim insertionRange As Range
    Dim fieldRange As Range
    Dim numberRange As Range
    Dim sequenceField As Field
    Dim existingHelperField As Field
    Dim visibleNumberField As Field
    Dim candidateField As Field
    Dim paragraphStart As Long
    Dim formulaStart As Long
    Dim insertionStart As Long
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim captionBookmarkName As String
    Dim targetBookmarkName As String
    Dim suffixText As String
    Dim operationStage As String
    Dim operationErrorNumber As Long
    Dim operationErrorDescription As String

    On Error GoTo NumberFailed
    If formulaShape Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7502, "VisualTeX", _
            "The numbered formula image target is invalid."
    End If

    Set documentObject = formulaShape.Range.Document
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    paragraphStart = paragraphRange.Start
    formulaStart = formulaShape.Range.Start
    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
    Set existingHelperField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)

    operationStage = "validate-image-paragraph"
    If paragraphRange.Information(wdWithInTable) Or _
       paragraphRange.InlineShapes.Count <> 1 Or _
       paragraphRange.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7502, "VisualTeX", _
            "The numbered image must occupy one ordinary Word paragraph."
    End If

    ' Migrate the r29 image layout without touching VisualTeX body references:
    ' remove only the old same-paragraph SEQ/REF tail, then rebuild the visible
    ' number as (REF VT_N_) while the true native SEQ lives in its own helper
    ' paragraph. This keeps VT_N_ as the single numbering source and makes
    ' Word's built-in Equation cross-reference list a pure number.
    operationStage = "validate-old-number-tail"
    Set suffixRange = documentObject.Range( _
        Start:=formulaShape.Range.End, End:=paragraphRange.End - 1)
    suffixText = suffixRange.Text
    For Each candidateField In suffixRange.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            suffixText = Replace$(suffixText, _
                candidateField.Result.Text, "", 1, 1, vbBinaryCompare)
        ElseIf candidateField.Type = wdFieldRef Then
            targetBookmarkName = VTReferenceTargetBookmarkName( _
                candidateField.Code.Text)
            If StrComp(targetBookmarkName, sequenceBookmarkName, _
               vbTextCompare) <> 0 Then
                Err.Raise vbObjectError + 7568, "VisualTeX", _
                    "The numbered image paragraph contains an unrelated REF field."
            End If
            suffixText = Replace$(suffixText, _
                candidateField.Result.Text, "", 1, 1, vbBinaryCompare)
        Else
            Err.Raise vbObjectError + 7568, "VisualTeX", _
                "The numbered image paragraph contains an unrelated field."
        End If
    Next candidateField
    suffixText = Replace$(suffixText, vbTab, "")
    suffixText = Replace$(suffixText, " ", "")
    suffixText = Replace$(suffixText, ChrW(160), "")
    suffixText = Replace$(suffixText, ChrW(8203), "")
    suffixText = Replace$(suffixText, ChrW(8288), "")
    suffixText = Replace$(suffixText, "(", "")
    suffixText = Replace$(suffixText, ")", "")
    If Len(suffixText) <> 0 Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The numbered image paragraph contains body text after the formula."
    End If

    operationStage = "remove-old-number-tail"
    If documentObject.Bookmarks.Exists(numberBookmarkName) Then
        documentObject.Bookmarks(numberBookmarkName).Delete
    End If
    If existingHelperField Is Nothing Then
        If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
            documentObject.Bookmarks(sequenceBookmarkName).Delete
        End If
        If documentObject.Bookmarks.Exists(captionBookmarkName) Then
            documentObject.Bookmarks(captionBookmarkName).Delete
        End If
    End If
    If suffixRange.End > suffixRange.Start Then suffixRange.Delete

    operationStage = "normalize-center-prefix"
    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, paragraphStart)
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    Set prefixRange = documentObject.Range( _
        Start:=paragraphRange.Start, End:=formulaShape.Range.Start)
    If VTWordRangeHasMeaningfulText(prefixRange) Then
        Err.Raise vbObjectError + 7497, "VisualTeX", _
            "A numbered image formula must occupy its own paragraph."
    End If
    If prefixRange.End > prefixRange.Start Then prefixRange.Delete
    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, paragraphStart)
    Set insertionRange = _
        VTPrependCenterTabPreservingImage(formulaShape.Range.Duplicate)
    Set formulaShape = insertionRange.InlineShapes(1)
    formulaStart = formulaShape.Range.Start
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    VTConfigureNumberedEquationParagraph paragraphRange

    operationStage = "create-external-sequence-helper"
    Set sequenceField = VTEnsureNativeEquationSequenceHelper( _
        formulaShape.Range.Duplicate, formulaId)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "Word did not create the image Equation SEQ helper paragraph."
    End If
    If Not VTHelperParagraphOwnsNativeEquationSequence( _
       sequenceField.Result.Paragraphs(1).Range.Duplicate) Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The image Equation SEQ helper paragraph is invalid."
    End If

    operationStage = "insert-visible-number-reference"
    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, paragraphStart)
    insertionStart = formulaShape.Range.End
    Set insertionRange = documentObject.Range( _
        Start:=insertionStart, End:=insertionStart)
    insertionRange.Text = vbTab & "()"
    Set fieldRange = documentObject.Range( _
        Start:=insertionStart + 2, End:=insertionStart + 2)
    Set visibleNumberField = documentObject.Fields.Add( _
        Range:=fieldRange, Type:=wdFieldRef, _
        Text:=VTParenthesizedEquationReferenceFieldText( _
            sequenceBookmarkName), _
        PreserveFormatting:=False)
    visibleNumberField.Update
    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, paragraphStart)
    Set visibleNumberField = VTImageEquationReferenceField( _
        formulaShape.Range.Duplicate, formulaId)
    If visibleNumberField Is Nothing Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "Word did not retain the image Equation visible REF."
    End If
    Set numberRange = VTImageEquationNumberRange( _
        formulaShape.Range.Duplicate, visibleNumberField)
    If numberRange Is Nothing Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "Word did not expose the image Equation visible number range."
    End If
    VTSetEquationNumberBookmarkExact _
        documentObject, formulaId, numberRange

    operationStage = "finalize-image-number"
    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, paragraphStart)
    VTConfigureNumberedEquationParagraph _
        formulaShape.Range.Paragraphs(1).Range.Duplicate
    VTFinalizeParagraphEquationNumber _
        documentObject, formulaShape.Range.Duplicate, formulaId, _
        deferReconcile
    Set formulaShape = VTResolveImageFormulaInParagraph( _
        documentObject, paragraphStart)
    Set VTInsertEquationNumber = _
        formulaShape.Range.Paragraphs(1).Range.Duplicate
    Exit Function

NumberFailed:
    operationErrorNumber = Err.Number
    operationErrorDescription = Err.Description
    Err.Raise operationErrorNumber, "VisualTeX Equation numbering", _
        "VTInsertEquationNumber/" & operationStage & ": " & _
        operationErrorDescription
End Function

Private Function VTEquationNumberRaisePoints( _
    ByVal formulaHeightPoints As Double, _
    ByVal numberFontSize As Single) As Single

    ' Establish a neutral baseline only. The final image-number position is
    ' measured from Word's real page coordinates and corrected afterwards by
    ' VTCalibrateImageEquationNumberPosition; no fixed offset is reliable.
    VTEquationNumberRaisePoints = 0!
End Function

Private Function VTNativeEquationNumberRaisePoints( _
    ByVal formulaRange As Range, _
    ByVal numberFontSize As Single) As Single

    ' Numbered OMML is a built-up inline OMath on an otherwise dedicated,
    ' center-tabbed paragraph. Word's line box already aligns the adjacent
    ' ordinary number text to that OMath; an extra Font.Position correction
    ' would create the same over-shift observed for numbered image formulas.
    VTNativeEquationNumberRaisePoints = 0!
End Function

Private Function VTNativeEquationLabelName() As String
    Dim equationLabel As CaptionLabel

    Set equationLabel = Application.CaptionLabels(wdCaptionEquation)
    VTNativeEquationLabelName = Trim$(equationLabel.Name)
    If Len(VTNativeEquationLabelName) = 0 Then
        Err.Raise vbObjectError + 7429, "VisualTeX", "Word did not expose its built-in Equation caption label."
    End If
End Function

Private Function VTEquationSequenceFieldText(ByVal equationLabelName As String) As String
    ' SEQ already produces decimal Arabic ordinals by default. Do not append
    ' the optional \* ARABIC switch: Word for Mac converts its ordinary * into
    ' the mathematical ∗ character when the field lives inside built-up OMath,
    ' which corrupts the field code and yields an undefined-Bookmark result.
    If InStr(1, equationLabelName, " ", vbBinaryCompare) > 0 Then
        VTEquationSequenceFieldText = _
            """" & Replace$(equationLabelName, """", """""") & """"
    Else
        VTEquationSequenceFieldText = equationLabelName
    End If
End Function

Private Function VTInsertRegisteredEquationCaption( _
    ByVal insertionRange As Range, _
    ByVal equationLabelName As String) As Field

    Dim documentObject As Document
    Dim sequenceField As Field
    Dim fieldParagraphRange As Range
    Dim insertionStart As Long
    Dim insertionParagraphStart As Long
    Dim captionStage As String
    Dim captionErrorNumber As Long
    Dim captionErrorDescription As String

    If insertionRange Is Nothing Then
        Err.Raise vbObjectError + 7430, "VisualTeX", _
            "The Equation caption insertion range is missing."
    End If
    Set documentObject = insertionRange.Document
    insertionStart = insertionRange.Start
    insertionParagraphStart = insertionRange.Paragraphs(1).Range.Start
    On Error GoTo CaptionFailed

    ' Match the validated Windows native-Office architecture without copying
    ' any Windows-only COM lifetime code: create a real SEQ field for Word's
    ' built-in, localized Equation caption label in a dedicated paragraph.
    ' Word recognizes this native SEQ field in GetCrossReferenceItems and
    ' InsertCrossReference; no synthetic number text is the sequence source.
    captionStage = "insert-native-seq"
    Set sequenceField = documentObject.Fields.Add( _
        Range:=insertionRange, _
        Type:=wdFieldEmpty, _
        Text:="SEQ " & VTEquationSequenceFieldText(equationLabelName), _
        PreserveFormatting:=True)
    sequenceField.Update

    captionStage = "verify-native-seq"
    If Not VTIsNativeEquationSequenceField( _
       sequenceField, equationLabelName) Then
        Err.Raise vbObjectError + 7425, "VisualTeX", _
            "Word did not create a native Equation SEQ field."
    End If
    Set fieldParagraphRange = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If fieldParagraphRange.Start <> insertionParagraphStart Or _
       Abs(VTEquationFieldStart(sequenceField) - insertionStart) > 16 Then
        Err.Raise vbObjectError + 7537, "VisualTeX", _
            "Word inserted the native Equation SEQ outside its helper paragraph."
    End If

    Set VTInsertRegisteredEquationCaption = sequenceField
    Exit Function

CaptionFailed:
    captionErrorNumber = Err.Number
    captionErrorDescription = Err.Description
    Err.Raise captionErrorNumber, "VisualTeX Equation caption", _
        "VTInsertRegisteredEquationCaption/" & captionStage & ": " & _
        captionErrorDescription
End Function

Private Function VTIsNativeEquationSequenceField( _
    ByVal candidate As Field, _
    ByVal equationLabelName As String) As Boolean

    Dim fieldCode As String
    If candidate Is Nothing Then Exit Function
    If candidate.Type <> wdFieldSequence Then Exit Function
    fieldCode = candidate.Code.Text
    VTIsNativeEquationSequenceField = _
        InStr(1, fieldCode, "SEQ", vbTextCompare) > 0 And _
        InStr(1, fieldCode, equationLabelName, vbTextCompare) > 0
End Function

Private Function VTEquationSequenceResultText( _
    ByVal sequenceField As Field) As String

    Dim resultText As String

    If sequenceField Is Nothing Then Exit Function
    resultText = sequenceField.Result.Text
    resultText = Replace$(resultText, vbCr, "")
    resultText = Replace$(resultText, Chr$(7), "")
    resultText = Replace$(resultText, Chr$(11), "")
    VTEquationSequenceResultText = Trim$(resultText)
End Function

Private Function VTEquationSequenceFieldCodeForOrdinal( _
    ByVal equationLabelName As String, _
    ByVal sequenceOrdinal As Long) As String

    If sequenceOrdinal < 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The Equation sequence ordinal must be positive."
    End If

    ' Keep every caption as a normal flowing Word SEQ field. A forced \r
    ' restart rewrites the field whenever an earlier formula is removed and
    ' invalidates Word's native cross-reference target. The field code also
    ' deliberately has no \* formatting switch so built-up OMath cannot
    ' transform its asterisk into a mathematical operator.
    VTEquationSequenceFieldCodeForOrdinal = _
        " SEQ " & VTEquationSequenceFieldText(equationLabelName) & " "
End Function

Private Function VTNormalizeEquationFieldCode( _
    ByVal fieldCode As String) As String

    Dim normalizedCode As String

    normalizedCode = Replace$(fieldCode, vbCr, " ")
    normalizedCode = Replace$(normalizedCode, vbLf, " ")
    normalizedCode = Replace$(normalizedCode, Chr$(11), " ")
    normalizedCode = Trim$(normalizedCode)
    Do While InStr(1, normalizedCode, "  ", vbBinaryCompare) > 0
        normalizedCode = Replace$(normalizedCode, "  ", " ")
    Loop
    VTNormalizeEquationFieldCode = normalizedCode
End Function

Private Function VTEquationSequenceFieldHasOrdinal( _
    ByVal sequenceField As Field, _
    ByVal equationLabelName As String, _
    ByVal sequenceOrdinal As Long) As Boolean

    Dim actualCode As String
    Dim expectedCode As String

    If sequenceField Is Nothing Or sequenceOrdinal < 1 Then Exit Function
    actualCode = VTNormalizeEquationFieldCode(sequenceField.Code.Text)
    expectedCode = VTNormalizeEquationFieldCode( _
        VTEquationSequenceFieldCodeForOrdinal( _
            equationLabelName, sequenceOrdinal))
    VTEquationSequenceFieldHasOrdinal = _
        StrComp(actualCode, expectedCode, vbTextCompare) = 0 And _
        VTEquationSequenceResultText(sequenceField) = CStr(sequenceOrdinal)
End Function

Private Sub VTApplyEquationSequenceOrdinal( _
    ByVal sequenceField As Field, _
    ByVal equationLabelName As String, _
    ByVal sequenceOrdinal As Long)

    Dim expectedCode As String

    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The Equation SEQ field is missing."
    End If
    expectedCode = VTEquationSequenceFieldCodeForOrdinal( _
        equationLabelName, sequenceOrdinal)

    ' Migrate the legacy restarted field once. After migration, renumbering
    ' updates only the field result, so Word can keep its own _Ref target.
    If StrComp(VTNormalizeEquationFieldCode(sequenceField.Code.Text), _
       VTNormalizeEquationFieldCode(expectedCode), vbTextCompare) <> 0 Then
        sequenceField.Code.Text = expectedCode
    End If
    sequenceField.Update

    If Not VTEquationSequenceFieldHasOrdinal( _
       sequenceField, equationLabelName, sequenceOrdinal) Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word did not update the flowing native Equation SEQ field" & _
            " [code=" & sequenceField.Code.Text & _
            "; result=" & sequenceField.Result.Text & _
            "; expected=" & CStr(sequenceOrdinal) & "]."
    End If
End Sub

Private Function VTEquationSequenceOrdinal( _
    ByVal documentObject As Document, _
    ByVal targetField As Field, _
    ByVal equationLabelName As String) As Long

    Dim candidate As Field
    Dim sequenceOrdinal As Long
    Dim targetStart As Long

    If documentObject Is Nothing Or targetField Is Nothing Then Exit Function
    targetStart = VTEquationFieldStart(targetField)
    For Each candidate In documentObject.Fields
        If VTIsNativeEquationSequenceField(candidate, equationLabelName) Then
            sequenceOrdinal = sequenceOrdinal + 1
            If VTEquationFieldStart(candidate) = targetStart Then
                VTEquationSequenceOrdinal = sequenceOrdinal
                Exit Function
            End If
        End If
    Next candidate
End Function

Private Function VTEquationFieldStart( _
    ByVal sequenceField As Field) As Long

    If sequenceField Is Nothing Or sequenceField.Code.Start <= 0 Then
        Err.Raise vbObjectError + 7534, "VisualTeX", _
            "The Equation field start boundary is invalid."
    End If
    VTEquationFieldStart = sequenceField.Code.Start - 1
End Function

Private Function VTEquationFieldEnd( _
    ByVal sequenceField As Field) As Long

    If sequenceField Is Nothing Or sequenceField.Result.End < 0 Then
        Err.Raise vbObjectError + 7535, "VisualTeX", _
            "The Equation field end boundary is invalid."
    End If
    VTEquationFieldEnd = sequenceField.Result.End + 1
End Function

Private Function VTResolveEquationSequenceFieldNear( _
    ByVal documentObject As Document, _
    ByVal expectedStart As Long, _
    ByVal maximumDistance As Long) As Field

    Dim candidate As Field
    Dim match As Field
    Dim probeRange As Range
    Dim equationLabelName As String
    Dim candidateStart As Long
    Dim candidateDistance As Long
    Dim probeStart As Long
    Dim probeEnd As Long
    Dim bestDistance As Long
    Dim matchCount As Long

    If documentObject Is Nothing Or expectedStart < 0 Or _
       maximumDistance < 0 Then
        Err.Raise vbObjectError + 7536, "VisualTeX", _
            "The Equation field resolver received an invalid target."
    End If
    equationLabelName = VTNativeEquationLabelName()
    probeStart = expectedStart - maximumDistance - 16
    If probeStart < 0 Then probeStart = 0
    probeEnd = expectedStart + maximumDistance + 32
    If probeEnd > documentObject.Content.End Then
        probeEnd = documentObject.Content.End
    End If
    Set probeRange = documentObject.Range( _
        Start:=probeStart, End:=probeEnd)
    bestDistance = 2147483647
    For Each candidate In probeRange.Fields
        If VTIsNativeEquationSequenceField(candidate, equationLabelName) Then
            candidateStart = VTEquationFieldStart(candidate)
            candidateDistance = Abs(candidateStart - expectedStart)
            If candidateDistance <= maximumDistance Then
                If candidateDistance < bestDistance Then
                    bestDistance = candidateDistance
                    matchCount = 1
                    Set match = candidate
                ElseIf candidateDistance = bestDistance Then
                    matchCount = matchCount + 1
                End If
            End If
        End If
    Next candidate
    If matchCount <> 1 Or match Is Nothing Then
        Err.Raise vbObjectError + 7536, "VisualTeX", _
            "Word could not re-resolve the Equation field after layout changes."
    End If
    Set VTResolveEquationSequenceFieldNear = match
End Function

Private Function VTResolveVisibleEquationReferenceFieldNear( _
    ByVal documentObject As Document, _
    ByVal expectedStart As Long, _
    ByVal sequenceBookmarkName As String, _
    ByVal maximumDistance As Long) As Field

    Dim candidate As Field
    Dim match As Field
    Dim probeRange As Range
    Dim targetBookmarkName As String
    Dim candidateStart As Long
    Dim candidateDistance As Long
    Dim probeStart As Long
    Dim probeEnd As Long
    Dim bestDistance As Long
    Dim matchCount As Long

    If documentObject Is Nothing Or expectedStart < 0 Or _
       Len(sequenceBookmarkName) = 0 Or maximumDistance < 0 Then
        Err.Raise vbObjectError + 7536, "VisualTeX", _
            "The Equation reference resolver received an invalid target."
    End If
    probeStart = expectedStart - maximumDistance - 16
    If probeStart < 0 Then probeStart = 0
    probeEnd = expectedStart + maximumDistance + 64
    If probeEnd > documentObject.Content.End Then
        probeEnd = documentObject.Content.End
    End If
    Set probeRange = documentObject.Range( _
        Start:=probeStart, End:=probeEnd)
    bestDistance = 2147483647

    For Each candidate In probeRange.Fields
        If candidate.Type = wdFieldRef And _
           candidate.Result.Information(wdWithInTable) Then
            targetBookmarkName = _
                VTReferenceTargetBookmarkName(candidate.Code.Text)
            If StrComp(targetBookmarkName, sequenceBookmarkName, _
               vbTextCompare) = 0 Then
                candidateStart = VTEquationFieldStart(candidate)
                candidateDistance = Abs(candidateStart - expectedStart)
                If candidateDistance <= maximumDistance Then
                    If candidateDistance < bestDistance Then
                        bestDistance = candidateDistance
                        matchCount = 1
                        Set match = candidate
                    ElseIf candidateDistance = bestDistance Then
                        matchCount = matchCount + 1
                    End If
                End If
            End If
        End If
    Next candidate

    If matchCount <> 1 Or match Is Nothing Then
        Err.Raise vbObjectError + 7536, "VisualTeX", _
            "Word could not re-resolve the Equation reference field after layout changes."
    End If
    Set VTResolveVisibleEquationReferenceFieldNear = match
End Function

Private Function VTEquationCrossReferenceText( _
    ByVal latexBase64 As String) As String

    Dim value As String

    If Not VTIsBase64UrlPayload(latexBase64) Then
        Err.Raise vbObjectError + 7491, "VisualTeX", _
            "The Equation cross-reference LaTeX payload is invalid."
    End If
    value = VTBase64UrlDecodeUtf8(latexBase64)
    value = Replace$(value, vbCr, " ")
    value = Replace$(value, vbLf, " ")
    value = Replace$(value, vbTab, " ")
    value = Replace$(value, ChrW(160), " ")
    Do While InStr(1, value, "  ", vbBinaryCompare) > 0
        value = Replace$(value, "  ", " ")
    Loop
    value = Trim$(value)
    If Len(value) = 0 Then value = "VisualTeX formula"
    If Len(value) > 240 Then value = Left$(value, 237) & "..."
    VTEquationCrossReferenceText = value
End Function

Private Function VTEquationCaptionBookmarkName( _
    ByVal formulaId As String) As String

    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7492, "VisualTeX", _
            "VisualTeX cannot bookmark Equation caption text with an invalid formula id."
    End If
    VTEquationCaptionBookmarkName = _
        VT_WORD_CAPTION_BOOKMARK_PREFIX & Replace$(formulaId, "-", "")
    If Len(VTEquationCaptionBookmarkName) > 40 Then
        Err.Raise vbObjectError + 7493, "VisualTeX", _
            "The Equation caption Bookmark name is longer than Word permits."
    End If
End Function

Private Sub VTSetCollapsedEquationCaptionBookmark( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal paragraphRange As Range)

    Dim bookmarkName As String
    Dim captionRange As Range

    If documentObject Is Nothing Or paragraphRange Is Nothing Then
        Err.Raise vbObjectError + 7567, "VisualTeX", _
            "The Equation caption identity target is missing."
    End If
    Set captionRange = paragraphRange.Duplicate
    If captionRange.End > captionRange.Start Then
        captionRange.End = captionRange.End - 1
    End If
    captionRange.Collapse wdCollapseEnd
    bookmarkName = VTEquationCaptionBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        documentObject.Bookmarks(bookmarkName).Delete
    End If
    documentObject.Bookmarks.Add _
        name:=bookmarkName, Range:=captionRange
End Sub

Private Function VTEquationCaptionBookmarkIsCollapsedInParagraph( _
    ByVal captionRange As Range, _
    ByVal paragraphRange As Range) As Boolean

    If captionRange Is Nothing Or paragraphRange Is Nothing Then Exit Function
    If captionRange.Start <> captionRange.End Then Exit Function
    If captionRange.Start < paragraphRange.Start Or _
       captionRange.Start >= paragraphRange.End Then Exit Function
    If captionRange.Information(wdWithInTable) <> _
       paragraphRange.Information(wdWithInTable) Then Exit Function
    VTEquationCaptionBookmarkIsCollapsedInParagraph = True
End Function

Private Function VTEquationNumberBookmarkName( _
    ByVal formulaId As String) As String

    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7546, "VisualTeX", _
            "VisualTeX cannot bookmark an Equation number with an invalid formula id."
    End If
    VTEquationNumberBookmarkName = _
        VT_WORD_NUMBER_BOOKMARK_PREFIX & Replace$(formulaId, "-", "")
    If Len(VTEquationNumberBookmarkName) > 40 Then
        Err.Raise vbObjectError + 7546, "VisualTeX", _
            "The Equation number Bookmark name is longer than Word permits."
    End If
End Function

Private Function VTEquationSequenceNumberBookmarkName( _
    ByVal formulaId As String) As String

    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7546, "VisualTeX", _
            "VisualTeX cannot bookmark an Equation sequence with an invalid formula id."
    End If
    VTEquationSequenceNumberBookmarkName = _
        VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX & Replace$(formulaId, "-", "")
    If Len(VTEquationSequenceNumberBookmarkName) > 40 Then
        Err.Raise vbObjectError + 7546, "VisualTeX", _
            "The Equation sequence Bookmark name is longer than Word permits."
    End If
End Function

Private Function VTParenthesizedEquationReferenceFieldText( _
    ByVal targetBookmarkName As String) As String

    If Len(Trim$(targetBookmarkName)) = 0 Then
        Err.Raise vbObjectError + 7546, "VisualTeX", _
            "The Equation reference Bookmark name is missing."
    End If
    ' The right cell points to the exact native VT_N_ SEQ result Bookmark.
    ' Parentheses remain ordinary text outside the REF field, matching Word's
    ' validated native numbering structure and avoiding a Bookmark that spans
    ' literal characters and a field boundary.
    VTParenthesizedEquationReferenceFieldText = _
        targetBookmarkName & " \h"
End Function

Private Function VTEquationSequenceFieldForBookmark( _
    ByVal documentObject As Document, _
    ByVal sequenceBookmarkName As String) As Field

    Dim candidate As Field
    Dim bookmarkRange As Range
    Dim match As Field
    Dim matchCount As Long
    Dim equationLabelName As String

    If documentObject Is Nothing Or Len(sequenceBookmarkName) = 0 Then
        Exit Function
    End If
    If Not documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        Exit Function
    End If
    Set bookmarkRange = documentObject.Bookmarks( _
        sequenceBookmarkName).Range.Duplicate
    equationLabelName = VTNativeEquationLabelName()
    For Each candidate In documentObject.Fields
        If VTIsNativeEquationSequenceField(candidate, equationLabelName) Then
            If candidate.Result.Start <= bookmarkRange.Start Then
                If candidate.Result.End >= bookmarkRange.End Then
                    matchCount = matchCount + 1
                    Set match = candidate
                End If
            End If
        End If
    Next candidate
    If matchCount = 1 Then
        Set VTEquationSequenceFieldForBookmark = match
    ElseIf matchCount > 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The Equation sequence Bookmark resolves to multiple fields."
    End If
End Function

Private Sub VTFormatHiddenEquationParagraph(ByVal paragraphRange As Range)
    Dim candidateField As Field
    Dim equationLabelName As String
    Dim visibleSize As Single

    If paragraphRange Is Nothing Then Exit Sub
    visibleSize = VTVisibleEquationNumberFontSize(paragraphRange.Document)
    equationLabelName = VTNativeEquationLabelName()
    With paragraphRange
        ' Keep the helper paragraph visually unobtrusive but not Word-hidden.
        ' Its exact one-point line box preserves the compact numbered layout.
        .Font.Size = 1!
        .Font.Hidden = False
        .Font.Color = wdColorWhite
        .Font.Position = 0
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            ' Keep the native SEQ result at normal visible formatting for Word's
            ' built-in cross-reference insertion, but position its compact helper
            ' paragraph five inches left of the text margin so no clipped glyph
            ' fragment is painted beside the display formula.
            .LeftIndent = -360!
            .RightIndent = 0!
            .FirstLineIndent = 0!
            .SpaceBefore = 0!
            .SpaceAfter = 0!
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 1!
        End With
    End With

    ' Word's built-in Cross-reference dialog copies direct character formatting
    ' from the SEQ result before VisualTeX has any callback opportunity. Keep the
    ' source result at normal body size and automatic color so the first native
    ' insertion is visible. The exact one-point paragraph line box, not character
    ' color, keeps the helper compact.
    For Each candidateField In paragraphRange.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, equationLabelName) Then
            With candidateField.Result.Font
                .Size = visibleSize
                .Hidden = False
                .Color = wdColorAutomatic
                .Position = 0
            End With
        End If
    Next candidateField
End Sub

Private Sub VTRefreshEquationNumberMirror( _
    ByVal documentObject As Document, _
    ByVal sequenceField As Field, _
    ByVal sequenceBookmarkName As String, _
    ByVal sequenceOrdinal As Long)

    Dim sequenceParagraph As Range
    Dim oldNumberRange As Range
    Dim oldNumberParagraph As Range
    Dim captionRange As Range
    Dim numberBookmarkName As String
    Dim captionBookmarkName As String
    Dim suffixText As String
    Dim expectedText As String
    Dim formulaId As String
    Dim fieldAnchor As Long
    Dim equationLabelName As String
    Dim paragraphNumber As Boolean

    If documentObject Is Nothing Or sequenceField Is Nothing Or _
       Len(sequenceBookmarkName) = 0 Or sequenceOrdinal < 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The native Equation number target is missing."
    End If
    suffixText = Mid$(sequenceBookmarkName, _
        Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX) + 1)
    numberBookmarkName = VT_WORD_NUMBER_BOOKMARK_PREFIX & suffixText
    captionBookmarkName = VT_WORD_CAPTION_BOOKMARK_PREFIX & suffixText

    ' The caller already resolved the exact SEQ field. Retain its local anchor
    ' instead of scanning every document field again for each formula.
    fieldAnchor = VTEquationFieldStart(sequenceField)
    equationLabelName = VTNativeEquationLabelName()
    VTApplyEquationSequenceOrdinal _
        sequenceField, equationLabelName, sequenceOrdinal
    Set sequenceField = VTResolveEquationSequenceFieldNear( _
        documentObject, fieldAnchor, 64)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The Equation number target lost its native SEQ field after update."
    End If
    If Not VTEquationSequenceFieldHasOrdinal( _
       sequenceField, equationLabelName, sequenceOrdinal) Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The native Equation SEQ field has an invalid ordinal code" & _
            " [code=" & sequenceField.Code.Text & _
            "; result=" & sequenceField.Result.Text & _
            "; expected=" & CStr(sequenceOrdinal) & _
            "; range=" & CStr(sequenceField.Result.Start) & "-" & _
                CStr(sequenceField.Result.End) & "]."
    End If
    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        documentObject.Bookmarks(sequenceBookmarkName).Delete
    End If
    documentObject.Bookmarks.Add _
        name:=sequenceBookmarkName, _
        Range:=sequenceField.Result.Duplicate

    formulaId = VTFormulaIdFromBookmarkSuffix(suffixText)
    Set sequenceParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    paragraphNumber = _
        (sequenceParagraph.InlineShapes.Count + _
         sequenceParagraph.OMaths.Count > 0)
    If Not paragraphNumber And _
       documentObject.Bookmarks.Exists(numberBookmarkName) Then
        Set oldNumberRange = documentObject.Bookmarks( _
            numberBookmarkName).Range.Duplicate
        paragraphNumber = Not oldNumberRange.Information(wdWithInTable)
    End If
    If paragraphNumber Then
        If Len(formulaId) = 0 Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The single-paragraph Equation number has no formula identity."
        End If
        VTRefreshParagraphEquationBookmarks _
            documentObject, sequenceField, formulaId
        Exit Sub
    End If

    ' Remove only legacy plain-text mirror artifacts. A current VT_R_ Bookmark
    ' lives in the visible right table cell and is rebuilt after the native REF
    ' field is inserted; never delete that table content here.
    If documentObject.Bookmarks.Exists(captionBookmarkName) Then
        documentObject.Bookmarks(captionBookmarkName).Delete
    End If
    If documentObject.Bookmarks.Exists(numberBookmarkName) Then
        Set oldNumberRange = documentObject.Bookmarks( _
            numberBookmarkName).Range.Duplicate
        documentObject.Bookmarks(numberBookmarkName).Delete
        If Not oldNumberRange.Information(wdWithInTable) Then
            Set oldNumberParagraph = _
                oldNumberRange.Paragraphs(1).Range.Duplicate
            If oldNumberParagraph.Fields.Count = 0 Then
                oldNumberParagraph.Delete
            ElseIf oldNumberRange.End > oldNumberRange.Start Then
                oldNumberRange.Delete
            End If
        End If
    End If

    Set sequenceField = VTResolveEquationSequenceFieldNear( _
        documentObject, fieldAnchor, 64)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The Equation native target cleanup lost its SEQ field."
    End If
    expectedText = CStr(sequenceOrdinal)
    If VTEquationSequenceResultText(sequenceField) <> expectedText Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word did not produce the native Equation SEQ result" & _
            " [code=" & sequenceField.Code.Text & _
            "; result=" & sequenceField.Result.Text & _
            "; expected=" & expectedText & "]."
    End If

    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        documentObject.Bookmarks(sequenceBookmarkName).Delete
    End If
    documentObject.Bookmarks.Add _
        name:=sequenceBookmarkName, _
        Range:=sequenceField.Result.Duplicate
    Set sequenceParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    VTFormatHiddenEquationParagraph sequenceParagraph
    Set captionRange = sequenceParagraph.Duplicate
    documentObject.Bookmarks.Add _
        name:=captionBookmarkName, Range:=captionRange
End Sub

Private Sub VTSetEquationNumberBookmark( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal openingRange As Range, _
    ByVal closingRange As Range)

    Dim numberRange As Range

    If documentObject Is Nothing Or openingRange Is Nothing Or _
       closingRange Is Nothing Then
        Err.Raise vbObjectError + 7546, "VisualTeX", _
            "The Equation number Bookmark target is missing."
    End If
    Set numberRange = documentObject.Range( _
        Start:=openingRange.Start, End:=closingRange.End)
    VTSetEquationNumberBookmarkExact documentObject, formulaId, numberRange
End Sub

Private Sub VTSetEquationNumberBookmarkExact( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal numberRange As Range)

    Dim bookmarkName As String

    If documentObject Is Nothing Or numberRange Is Nothing Then
        Err.Raise vbObjectError + 7546, "VisualTeX", _
            "The exact Equation number Bookmark target is missing."
    End If
    If numberRange.End <= numberRange.Start Then
        Err.Raise vbObjectError + 7546, "VisualTeX", _
            "The exact Equation number Bookmark target is empty."
    End If
    bookmarkName = VTEquationNumberBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        documentObject.Bookmarks(bookmarkName).Delete
    End If
    documentObject.Bookmarks.Add name:=bookmarkName, Range:=numberRange
End Sub

Private Function VTNativeEquationArrayMarkerRange( _
    ByVal formulaRange As Range) As Range

    Dim markerRange As Range

    If formulaRange Is Nothing Then Exit Function
    If formulaRange.OMaths.Count <> 1 Then Exit Function
    Set markerRange = formulaRange.OMaths(1).Range.Duplicate
    With markerRange.Find
        .ClearFormatting
        .Text = "#"
        .Forward = False
        .Wrap = wdFindStop
        .Format = False
    End With
    If markerRange.Find.Execute Then
        Set VTNativeEquationArrayMarkerRange = markerRange.Duplicate
    End If
End Function

Private Function VTNativeEquationFormulaContentRange( _
    ByVal formulaRange As Range) As Range

    Dim exactRange As Range
    Dim markerRange As Range

    If formulaRange Is Nothing Then Exit Function
    If formulaRange.OMaths.Count <> 1 Then Exit Function
    Set exactRange = formulaRange.OMaths(1).Range.Duplicate
    Set markerRange = VTNativeEquationArrayMarkerRange(exactRange)
    If Not markerRange Is Nothing Then exactRange.End = markerRange.Start
    If exactRange.End <= exactRange.Start Then Exit Function
    Set VTNativeEquationFormulaContentRange = exactRange
End Function

Private Function VTNativeEquationSequenceIsInsideMath( _
    ByVal formulaRange As Range, _
    ByVal sequenceField As Field) As Boolean

    Dim exactRange As Range
    Dim fieldStart As Long
    Dim fieldEnd As Long

    If formulaRange Is Nothing Or sequenceField Is Nothing Then Exit Function
    If formulaRange.OMaths.Count <> 1 Then Exit Function
    If Not VTIsNativeEquationSequenceField( _
       sequenceField, VTNativeEquationLabelName()) Then Exit Function
    Set exactRange = formulaRange.OMaths(1).Range.Duplicate
    fieldStart = VTEquationFieldStart(sequenceField)
    fieldEnd = VTEquationFieldEnd(sequenceField)
    VTNativeEquationSequenceIsInsideMath = _
        fieldStart >= exactRange.Start And fieldEnd <= exactRange.End
End Function

Private Function VTImageEquationReferenceField( _
    ByVal formulaRange As Range, _
    ByVal formulaId As String) As Field

    Dim exactRange As Range
    Dim paragraphRange As Range
    Dim candidateField As Field
    Dim targetBookmarkName As String
    Dim candidateTargetName As String
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim matchCount As Long
    Dim match As Field

    If formulaRange Is Nothing Then Exit Function
    If formulaRange.InlineShapes.Count <> 1 Or _
       formulaRange.OMaths.Count <> 0 Or _
       Not VTIsCanonicalUuid(formulaId) Then Exit Function
    Set exactRange = formulaRange.InlineShapes(1).Range.Duplicate
    Set paragraphRange = exactRange.Paragraphs(1).Range.Duplicate
    targetBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    For Each candidateField In paragraphRange.Fields
        If candidateField.Type = wdFieldRef Then
            fieldStart = VTEquationFieldStart(candidateField)
            fieldEnd = VTEquationFieldEnd(candidateField)
            If fieldStart > exactRange.End And _
               fieldEnd < paragraphRange.End Then
                candidateTargetName = VTReferenceTargetBookmarkName( _
                    candidateField.Code.Text)
                If StrComp(candidateTargetName, targetBookmarkName, _
                   vbTextCompare) = 0 Then
                    matchCount = matchCount + 1
                    Set match = candidateField
                End If
            End If
        End If
    Next candidateField
    If matchCount > 1 Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The image Equation paragraph contains multiple visible number references."
    End If
    If matchCount = 1 Then Set VTImageEquationReferenceField = match
End Function

Private Function VTImageEquationNumberRange( _
    ByVal formulaRange As Range, _
    ByVal numberField As Field) As Range

    Dim exactRange As Range
    Dim paragraphRange As Range
    Dim fieldStart As Long
    Dim fieldEnd As Long

    If formulaRange Is Nothing Or numberField Is Nothing Then Exit Function
    If formulaRange.InlineShapes.Count <> 1 Or _
       formulaRange.OMaths.Count <> 0 Or _
       numberField.Type <> wdFieldRef Then Exit Function
    Set exactRange = formulaRange.InlineShapes(1).Range.Duplicate
    Set paragraphRange = exactRange.Paragraphs(1).Range.Duplicate
    fieldStart = VTEquationFieldStart(numberField)
    fieldEnd = VTEquationFieldEnd(numberField)
    If fieldStart <= exactRange.End Or _
       fieldStart < 2 Or fieldEnd >= paragraphRange.End Or _
       exactRange.Document.Range(fieldStart - 2, fieldStart - 1).Text <> vbTab Or _
       exactRange.Document.Range(fieldStart - 1, fieldStart).Text <> "(" Or _
       exactRange.Document.Range(fieldEnd, fieldEnd + 1).Text <> ")" Then
        Exit Function
    End If
    Set VTImageEquationNumberRange = exactRange.Document.Range( _
        Start:=fieldStart - 1, End:=fieldEnd + 1)
End Function

Private Function VTNativeEquationArrayReferenceField( _
    ByVal formulaRange As Range, _
    ByVal formulaId As String) As Field

    Dim exactRange As Range
    Dim candidateField As Field
    Dim targetBookmarkName As String
    Dim candidateTargetName As String
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim matchCount As Long
    Dim match As Field

    If formulaRange Is Nothing Then Exit Function
    If formulaRange.OMaths.Count <> 1 Or _
       Not VTIsCanonicalUuid(formulaId) Then Exit Function
    Set exactRange = formulaRange.OMaths(1).Range.Duplicate
    targetBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    For Each candidateField In exactRange.Document.Fields
        If candidateField.Type = wdFieldRef Then
            fieldStart = VTEquationFieldStart(candidateField)
            fieldEnd = VTEquationFieldEnd(candidateField)
            If fieldStart >= exactRange.Start And fieldEnd <= exactRange.End Then
                candidateTargetName = VTReferenceTargetBookmarkName( _
                    candidateField.Code.Text)
                If StrComp(candidateTargetName, targetBookmarkName, _
                   vbTextCompare) = 0 Then
                    matchCount = matchCount + 1
                    Set match = candidateField
                End If
            End If
        End If
    Next candidateField
    If matchCount > 1 Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The native Equation array contains multiple visible number references."
    End If
    If matchCount = 1 Then Set VTNativeEquationArrayReferenceField = match
End Function

Private Function VTNativeEquationNumberIsInsideMath( _
    ByVal formulaRange As Range, _
    ByVal numberField As Field) As Boolean

    Dim exactRange As Range
    Dim markerRange As Range
    Dim targetBookmarkName As String
    Dim fieldStart As Long
    Dim fieldEnd As Long

    If formulaRange Is Nothing Or numberField Is Nothing Then Exit Function
    If formulaRange.OMaths.Count <> 1 Then Exit Function
    If numberField.Type <> wdFieldRef Then Exit Function
    Set exactRange = formulaRange.OMaths(1).Range.Duplicate
    Set markerRange = VTNativeEquationArrayMarkerRange(exactRange)
    If markerRange Is Nothing Then Exit Function
    targetBookmarkName = VTReferenceTargetBookmarkName(numberField.Code.Text)
    If Left$(targetBookmarkName, _
       Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)) <> _
       VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX Then Exit Function
    fieldStart = VTEquationFieldStart(numberField)
    fieldEnd = VTEquationFieldEnd(numberField)
    VTNativeEquationNumberIsInsideMath = _
        markerRange.Start > exactRange.Start And _
        fieldStart > markerRange.Start And fieldEnd < exactRange.End
End Function

Private Function VTNativeEquationArrayNumberRange( _
    ByVal formulaRange As Range, _
    ByVal numberField As Field) As Range

    Dim exactRange As Range
    Dim markerRange As Range
    Dim fieldStart As Long
    Dim fieldEnd As Long

    If formulaRange Is Nothing Or numberField Is Nothing Then Exit Function
    If formulaRange.OMaths.Count <> 1 Then Exit Function
    Set exactRange = formulaRange.OMaths(1).Range.Duplicate
    Set markerRange = VTNativeEquationArrayMarkerRange(exactRange)
    If markerRange Is Nothing Then Exit Function
    fieldStart = VTEquationFieldStart(numberField)
    fieldEnd = VTEquationFieldEnd(numberField)
    If markerRange.Start <= exactRange.Start Or _
       fieldStart <= markerRange.Start Or _
       fieldEnd >= exactRange.End Then Exit Function

    ' Word exposes built-up delimiter objects as OMath-internal boundaries. The
    ' durable visible-number region is the complete array tail after #; it must
    ' contain only the REF field that mirrors the external native SEQ result.
    Set VTNativeEquationArrayNumberRange = exactRange.Document.Range( _
        Start:=markerRange.End, End:=exactRange.End)
End Function

Private Function VTNativeEquationNumberBookmarkIsCompatible( _
    ByVal bookmarkRange As Range, _
    ByVal formulaRange As Range, _
    ByVal numberField As Field) As Boolean

    Dim exactRange As Range
    Dim expectedNumberRange As Range
    Dim paragraphRange As Range
    Dim beforeRange As Range
    Dim afterRange As Range
    Dim fieldStart As Long
    Dim fieldEnd As Long

    If bookmarkRange Is Nothing Or formulaRange Is Nothing Or _
       numberField Is Nothing Then Exit Function
    If formulaRange.OMaths.Count <> 1 Then Exit Function
    Set exactRange = formulaRange.OMaths(1).Range.Duplicate
    Set expectedNumberRange = VTNativeEquationArrayNumberRange( _
        exactRange, numberField)
    If expectedNumberRange Is Nothing Then Exit Function

    fieldStart = VTEquationFieldStart(numberField)
    fieldEnd = VTEquationFieldEnd(numberField)
    If bookmarkRange.Information(wdWithInTable) Or _
       fieldStart < bookmarkRange.Start Or _
       fieldEnd > bookmarkRange.End Then Exit Function

    If bookmarkRange.Start = expectedNumberRange.Start And _
       bookmarkRange.End = expectedNumberRange.End Then
        VTNativeEquationNumberBookmarkIsCompatible = True
        Exit Function
    End If
    If bookmarkRange.InlineShapes.Count <> 0 Or _
       bookmarkRange.OMaths.Count <> 1 Or _
       bookmarkRange.Start > exactRange.Start Or _
       bookmarkRange.End < exactRange.End Then Exit Function

    Set paragraphRange = exactRange.Paragraphs(1).Range.Duplicate
    If bookmarkRange.Start < paragraphRange.Start Or _
       bookmarkRange.End > paragraphRange.End Then Exit Function
    Set beforeRange = exactRange.Document.Range( _
        Start:=bookmarkRange.Start, End:=exactRange.Start)
    Set afterRange = exactRange.Document.Range( _
        Start:=exactRange.End, End:=bookmarkRange.End)
    If VTWordRangeHasMeaningfulText(beforeRange) Or _
       VTWordRangeHasMeaningfulText(afterRange) Then Exit Function

    VTNativeEquationNumberBookmarkIsCompatible = True
End Function

Private Function VTNativeEquationSequenceHelperField( _
    ByVal documentObject As Document, _
    ByVal formulaId As String) As Field

    Dim sequenceField As Field
    Dim helperParagraph As Range
    Dim sequenceBookmarkName As String

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then
        Exit Function
    End If
    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    Set sequenceField = VTEquationSequenceFieldForBookmark( _
        documentObject, sequenceBookmarkName)
    If sequenceField Is Nothing Then Exit Function
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If helperParagraph.Information(wdWithInTable) Or _
       helperParagraph.OMaths.Count <> 0 Or _
       helperParagraph.InlineShapes.Count <> 0 Or _
       helperParagraph.Fields.Count <> 1 Or _
       Not VTHelperParagraphOwnsNativeEquationSequence( _
           helperParagraph) Then Exit Function
    Set VTNativeEquationSequenceHelperField = sequenceField
End Function

Private Function VTEnsureNativeEquationSequenceHelper( _
    ByVal formulaRange As Range, _
    ByVal formulaId As String) As Field

    Dim documentObject As Document
    Dim exactFormulaRange As Range
    Dim formulaParagraph As Range
    Dim helperParagraph As Range
    Dim insertionRange As Range
    Dim fieldRange As Range
    Dim sequenceField As Field
    Dim candidateField As Field
    Dim sequenceBookmarkName As String
    Dim captionBookmarkName As String
    Dim helperStart As Long

    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The Equation sequence helper target is missing."
    End If
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The Equation sequence helper identity is invalid."
    End If
    If formulaRange.InlineShapes.Count = 1 And _
       formulaRange.OMaths.Count = 0 Then
        Set exactFormulaRange = _
            formulaRange.InlineShapes(1).Range.Duplicate
    ElseIf formulaRange.InlineShapes.Count = 0 And _
           formulaRange.OMaths.Count = 1 Then
        Set exactFormulaRange = formulaRange.OMaths(1).Range.Duplicate
    Else
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The Equation sequence helper target is ambiguous."
    End If
    Set documentObject = exactFormulaRange.Document
    Set formulaParagraph = _
        VTWordParagraphContainingFormula(exactFormulaRange)
    If formulaParagraph Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word could not resolve the native Equation paragraph for its sequence helper."
    End If

    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    If Not sequenceField Is Nothing Then
        Set helperParagraph = _
            sequenceField.Result.Paragraphs(1).Range.Duplicate
        If helperParagraph.Start < formulaParagraph.End Or _
           sequenceField.Result.OMaths.Count <> 0 Then
            Err.Raise vbObjectError + 7563, "VisualTeX", _
                "The native Equation SEQ is not isolated after the formula."
        End If
        For Each candidateField In exactFormulaRange.Fields
            If VTIsNativeEquationSequenceField( _
               candidateField, VTNativeEquationLabelName()) Then
                Err.Raise vbObjectError + 7563, "VisualTeX", _
                    "A native Equation SEQ was absorbed into OMath."
            End If
        Next candidateField
        captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
        VTSetCollapsedEquationCaptionBookmark _
            documentObject, formulaId, helperParagraph
        VTFormatHiddenEquationParagraph helperParagraph
        Set VTEnsureNativeEquationSequenceHelper = sequenceField
        Exit Function
    End If

    helperStart = formulaParagraph.End
    Set insertionRange = formulaParagraph.Duplicate
    insertionRange.InsertParagraphAfter
    Set helperParagraph = documentObject.Range( _
        Start:=helperStart, End:=helperStart).Paragraphs(1).Range.Duplicate
    If helperParagraph.Start <> helperStart Or _
       helperParagraph.Information(wdWithInTable) Or _
       helperParagraph.Fields.Count <> 0 Or _
       helperParagraph.InlineShapes.Count <> 0 Or _
       helperParagraph.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(helperParagraph) Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word did not create an independent native Equation SEQ paragraph."
    End If

    Set fieldRange = documentObject.Range( _
        Start:=helperStart, End:=helperStart)
    Set sequenceField = VTInsertRegisteredEquationCaption( _
        fieldRange, VTNativeEquationLabelName())
    sequenceField.Update
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If helperParagraph.Start <> helperStart Or _
       helperParagraph.Information(wdWithInTable) Or _
       helperParagraph.OMaths.Count <> 0 Or _
       helperParagraph.InlineShapes.Count <> 0 Or _
       helperParagraph.Fields.Count <> 1 Or _
       sequenceField.Result.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word absorbed the native Equation SEQ into formula math."
    End If
    For Each candidateField In exactFormulaRange.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            Err.Raise vbObjectError + 7563, "VisualTeX", _
                "The native Equation formula contains an unexpected SEQ field."
        End If
    Next candidateField

    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        documentObject.Bookmarks(sequenceBookmarkName).Delete
    End If
    documentObject.Bookmarks.Add _
        name:=sequenceBookmarkName, Range:=sequenceField.Result.Duplicate
    VTSetCollapsedEquationCaptionBookmark _
        documentObject, formulaId, helperParagraph
    VTFormatHiddenEquationParagraph helperParagraph
    Set VTEnsureNativeEquationSequenceHelper = sequenceField
End Function

Private Sub VTConfigureNativeEquationArrayParagraph( _
    ByVal paragraphRange As Range)

    If paragraphRange Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The native Equation array paragraph is missing."
    End If
    If paragraphRange.Style <> wdStyleCaption Then
        paragraphRange.Style = wdStyleCaption
    End If
    With paragraphRange.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .LeftIndent = 0!
        .RightIndent = 0!
        .FirstLineIndent = 0!
        .SpaceBefore = 0!
        .SpaceAfter = 0!
        .LineSpacingRule = wdLineSpaceSingle
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .WidowControl = True
        .TabStops.ClearAll
    End With
End Sub

Private Function VTCalibrateImageEquationNumberPosition( _
    ByVal formulaShape As InlineShape, _
    ByVal numberRange As Range) As Long

    Dim documentObject As Document
    Dim formulaProbe As Range
    Dim numberProbe As Range
    Dim formulaY As Single
    Dim numberY As Single
    Dim formulaCenterY As Single
    Dim numberCenterY As Single
    Dim numberLineHeight As Single
    Dim residual As Single
    Dim currentPosition As Long
    Dim nextPosition As Long
    Dim passIndex As Long

    If formulaShape Is Nothing Or numberRange Is Nothing Then
        Err.Raise vbObjectError + 7564, "VisualTeX", _
            "The image Equation visual-center calibration target is missing."
    End If
    Set documentObject = formulaShape.Range.Document
    numberRange.Font.Position = 0

    For passIndex = 1 To 4
        documentObject.Repaginate
        Set formulaProbe = formulaShape.Range.Duplicate
        formulaProbe.Collapse wdCollapseStart
        Set numberProbe = numberRange.Duplicate
        numberProbe.Collapse wdCollapseStart
        formulaY = CSng(formulaProbe.Information( _
            wdVerticalPositionRelativeToPage))
        numberY = CSng(numberProbe.Information( _
            wdVerticalPositionRelativeToPage))
        If formulaY < 0! Or numberY < 0! Then
            Err.Raise vbObjectError + 7564, "VisualTeX", _
                "Word did not expose image Equation vertical coordinates."
        End If

        numberLineHeight = numberRange.ParagraphFormat.LineSpacing
        If numberLineHeight <= 0! Or numberLineHeight = wdUndefined Or _
           numberLineHeight > 72! Then
            numberLineHeight = numberRange.Font.Size * 1.2!
        End If
        If numberLineHeight <= 0! Or numberLineHeight > 72! Then
            numberLineHeight = 14.4!
        End If
        currentPosition = numberRange.Font.Position
        If currentPosition = wdUndefined Then currentPosition = 0
        ' For an InlineShape, Word reports the text baseline/bottom anchor, not
        ' the bitmap top. For ordinary number text it likewise reports the line
        ' baseline before Font.Position is applied. Move upward by half-height
        ' from each baseline to compare the two visible centers.
        formulaCenterY = formulaY - formulaShape.Height / 2!
        numberCenterY = numberY - numberLineHeight / 2! - currentPosition
        residual = formulaCenterY - numberCenterY
        If Abs(residual) <= 1! Then Exit For

        ' Page coordinates increase downward and positive Font.Position raises
        ' the number. Preserve the measured feedback sign: a negative residual
        ' therefore produces the required positive raise for a tall image.
        nextPosition = currentPosition - CLng(residual)
        If nextPosition < -48 Then nextPosition = -48
        If nextPosition > 48 Then nextPosition = 48
        If nextPosition = currentPosition Then Exit For
        numberRange.Font.Position = nextPosition
    Next passIndex

    documentObject.Repaginate
    Set formulaProbe = formulaShape.Range.Duplicate
    formulaProbe.Collapse wdCollapseStart
    Set numberProbe = numberRange.Duplicate
    numberProbe.Collapse wdCollapseStart
    formulaY = CSng(formulaProbe.Information( _
        wdVerticalPositionRelativeToPage))
    numberY = CSng(numberProbe.Information( _
        wdVerticalPositionRelativeToPage))
    numberLineHeight = numberRange.ParagraphFormat.LineSpacing
    If numberLineHeight <= 0! Or numberLineHeight = wdUndefined Or _
       numberLineHeight > 72! Then
        numberLineHeight = numberRange.Font.Size * 1.2!
    End If
    If numberLineHeight <= 0! Or numberLineHeight > 72! Then
        numberLineHeight = 14.4!
    End If
    currentPosition = numberRange.Font.Position
    If currentPosition = wdUndefined Then currentPosition = 0
    formulaCenterY = formulaY - formulaShape.Height / 2!
    numberCenterY = numberY - numberLineHeight / 2! - currentPosition
    If formulaY < 0! Or numberY < 0! Or _
       Abs(formulaCenterY - numberCenterY) > 2! Then
        Err.Raise vbObjectError + 7564, "VisualTeX", _
            "Word could not align the image Equation and number visual centers" & _
            " [formulaBaseline=" & CStr(formulaY) & _
            "; formulaHeight=" & CStr(formulaShape.Height) & _
            "; formulaCenter=" & CStr(formulaCenterY) & _
            "; numberBaseline=" & CStr(numberY) & _
            "; numberLineHeight=" & CStr(numberLineHeight) & _
            "; numberCenter=" & CStr(numberCenterY) & _
            "; position=" & CStr(currentPosition) & "]."
    End If
    VTCalibrateImageEquationNumberPosition = numberRange.Font.Position
End Function

Private Sub VTRefreshParagraphEquationBookmarks( _
    ByVal documentObject As Document, _
    ByVal sequenceField As Field, _
    ByVal formulaId As String)

    Dim formulaParagraph As Range
    Dim sequenceParagraph As Range
    Dim numberRange As Range
    Dim formulaRange As Range
    Dim captionRange As Range
    Dim visibleNumberField As Field
    Dim candidateField As Field
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim captionBookmarkName As String
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim numberSize As Single
    Dim numberPosition As Long
    Dim beforeText As String
    Dim afterText As String

    If documentObject Is Nothing Or sequenceField Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The Equation identity target is missing."
    End If
    If sequenceField.Result.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "A table-free Equation sequence cannot be inside a table."
    End If
    Set sequenceParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate

    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        documentObject.Bookmarks(sequenceBookmarkName).Delete
    End If
    documentObject.Bookmarks.Add _
        name:=sequenceBookmarkName, Range:=sequenceField.Result.Duplicate

    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "Word could not resolve the visible formula from its number Bookmark."
    End If
    Set formulaParagraph = _
        VTWordParagraphContainingFormula(formulaRange)
    If formulaParagraph Is Nothing Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "Word could not resolve the visible Equation paragraph."
    End If
    If formulaParagraph.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The visible Equation is not in one ordinary Word paragraph."
    End If

    If formulaRange.InlineShapes.Count = 1 And _
       formulaRange.OMaths.Count = 0 Then
        If Not VTHelperParagraphOwnsNativeEquationSequence( _
           sequenceParagraph) Or _
           sequenceParagraph.Start < formulaParagraph.End Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The image Equation SEQ is not isolated after the formula."
        End If
        For Each candidateField In formulaParagraph.Fields
            If VTIsNativeEquationSequenceField( _
               candidateField, VTNativeEquationLabelName()) Then
                Err.Raise vbObjectError + 7560, "VisualTeX", _
                    "The image Equation paragraph still contains a native SEQ field."
            End If
        Next candidateField
        VTConfigureNumberedEquationParagraph formulaParagraph
        Set visibleNumberField = VTImageEquationReferenceField( _
            formulaRange, formulaId)
        If visibleNumberField Is Nothing Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The image Equation paragraph has no visible number REF."
        End If
        visibleNumberField.Update
        If VTEquationSequenceResultText(visibleNumberField) <> _
           VTEquationSequenceResultText(sequenceField) Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The visible image Equation REF does not match its external SEQ."
        End If
        Set numberRange = VTImageEquationNumberRange( _
            formulaRange, visibleNumberField)
        If numberRange Is Nothing Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "Word did not expose the image Equation number range."
        End If
        numberSize = visibleNumberField.Result.Font.Size
        numberPosition = 0
        VTFormatHiddenEquationParagraph sequenceParagraph
        Set captionRange = sequenceParagraph.Duplicate
        If captionRange.End > captionRange.Start Then
            captionRange.End = captionRange.End - 1
        End If
        captionRange.Collapse wdCollapseEnd
    ElseIf formulaRange.InlineShapes.Count = 0 And _
           formulaRange.OMaths.Count = 1 Then
        If Not VTHelperParagraphOwnsNativeEquationSequence( _
           sequenceParagraph) Or _
           sequenceParagraph.Start < formulaParagraph.End Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The native Equation SEQ is not isolated after the formula."
        End If
        For Each candidateField In formulaRange.OMaths(1).Range.Fields
            If VTIsNativeEquationSequenceField( _
               candidateField, VTNativeEquationLabelName()) Then
                Err.Raise vbObjectError + 7560, "VisualTeX", _
                    "The native Equation SEQ was absorbed into OMath."
            End If
        Next candidateField
        formulaRange.OMaths(1).Type = wdOMathDisplay
        formulaRange.OMaths(1).Justification = wdOMathJcCenterGroup
        VTConfigureNativeEquationArrayParagraph formulaParagraph
        Set visibleNumberField = VTNativeEquationArrayReferenceField( _
            formulaRange, formulaId)
        If visibleNumberField Is Nothing Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The native Equation array has no internal number REF."
        End If
        If Not VTNativeEquationNumberIsInsideMath( _
           formulaRange, visibleNumberField) Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The native Equation number REF left its OMath array."
        End If
        visibleNumberField.Update
        If VTEquationSequenceResultText(visibleNumberField) <> _
           VTEquationSequenceResultText(sequenceField) Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The visible native Equation REF does not match its external SEQ."
        End If
        Set numberRange = VTNativeEquationArrayNumberRange( _
            formulaRange, visibleNumberField)
        If numberRange Is Nothing Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "Word did not expose the native Equation array number range."
        End If
        numberSize = visibleNumberField.Result.Font.Size
        numberPosition = 0
        VTFormatHiddenEquationParagraph sequenceParagraph
        Set captionRange = sequenceParagraph.Duplicate
        If captionRange.End > captionRange.Start Then
            captionRange.End = captionRange.End - 1
        End If
        captionRange.Collapse wdCollapseEnd
    Else
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The Equation does not contain exactly one supported visible formula."
    End If

    If numberSize <= 0! Or numberSize > 72! Then
        numberSize = numberRange.Font.Size
    End If
    If numberSize <= 0! Or numberSize > 72! Then
        numberSize = VTVisibleEquationNumberFontSize(documentObject)
    End If
    If numberPosition = wdUndefined Or numberPosition < -48 Or _
       numberPosition > 48 Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The Equation number has an invalid vertical position."
    End If

    VTSetEquationNumberBookmarkExact _
        documentObject, formulaId, numberRange
    With numberRange.Font
        .Hidden = False
        .Color = wdColorAutomatic
        .Position = numberPosition
        .Size = numberSize
    End With
    If formulaRange.InlineShapes.Count = 1 Then
        numberPosition = VTCalibrateImageEquationNumberPosition( _
            formulaRange.InlineShapes(1), numberRange)
    End If

    If documentObject.Bookmarks.Exists(captionBookmarkName) Then
        documentObject.Bookmarks(captionBookmarkName).Delete
    End If
    documentObject.Bookmarks.Add _
        name:=captionBookmarkName, Range:=captionRange
End Sub

Private Sub VTVerifyParagraphEquationNumberIntegrity( _
    ByVal formulaRange As Range, _
    ByVal formulaId As String, _
    ByVal expectedOrdinal As Long)

    Dim documentObject As Document
    Dim formulaParagraph As Range
    Dim helperParagraph As Range
    Dim sequenceField As Field
    Dim visibleNumberField As Field
    Dim candidateField As Field
    Dim numberRange As Range
    Dim expectedNumberRange As Range
    Dim captionRange As Range
    Dim sequenceBookmarkName As String
    Dim captionBookmarkName As String
    Dim numberBookmarkName As String
    Dim expectedText As String

    If formulaRange Is Nothing Or expectedOrdinal < 1 Or _
       Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The Equation verification target is invalid."
    End If
    Set documentObject = formulaRange.Document
    Set formulaParagraph = _
        VTWordParagraphContainingFormula(formulaRange)
    If formulaParagraph Is Nothing Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "Word could not resolve the numbered formula paragraph."
    End If
    If formulaParagraph.Information(wdWithInTable) Or _
       formulaParagraph.Paragraphs.Count <> 1 Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The numbered formula is not one ordinary Word paragraph."
    End If

    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(sequenceBookmarkName) Or _
       Not documentObject.Bookmarks.Exists(captionBookmarkName) Or _
       Not documentObject.Bookmarks.Exists(numberBookmarkName) Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "Word did not preserve the Equation Bookmark set."
    End If
    Set sequenceField = VTEquationSequenceFieldForBookmark( _
        documentObject, sequenceBookmarkName)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The numbered formula has no native Equation SEQ field."
    End If
    expectedText = CStr(expectedOrdinal)
    If VTEquationSequenceResultText(sequenceField) <> expectedText Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The Equation SEQ result is incorrect."
    End If
    Set numberRange = documentObject.Bookmarks( _
        numberBookmarkName).Range.Duplicate
    Set captionRange = documentObject.Bookmarks( _
        captionBookmarkName).Range.Duplicate

    If formulaRange.InlineShapes.Count = 1 And _
       formulaRange.OMaths.Count = 0 Then
        Set helperParagraph = _
            sequenceField.Result.Paragraphs(1).Range.Duplicate
        Set visibleNumberField = VTImageEquationReferenceField( _
            formulaRange, formulaId)
        Set expectedNumberRange = VTImageEquationNumberRange( _
            formulaRange, visibleNumberField)
        If visibleNumberField Is Nothing Or _
           expectedNumberRange Is Nothing Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The image Equation visible REF range is missing."
        End If
        If Not VTHelperParagraphOwnsNativeEquationSequence( _
               helperParagraph) Or _
           helperParagraph.Start < formulaParagraph.End Or _
           VTEquationSequenceResultText(visibleNumberField) <> expectedText Or _
           numberRange.Start <> expectedNumberRange.Start Or _
           numberRange.End <> expectedNumberRange.End Or _
           numberRange.Information(wdWithInTable) Or _
           numberRange.Paragraphs(1).Range.Start <> _
               formulaParagraph.Start Or _
           numberRange.Text <> "(" & expectedText & ")" Or _
           Not VTEquationCaptionBookmarkIsCollapsedInParagraph( _
               captionRange, helperParagraph) Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The image Equation external SEQ or visible REF is incomplete."
        End If
        For Each candidateField In formulaParagraph.Fields
            If VTIsNativeEquationSequenceField( _
               candidateField, VTNativeEquationLabelName()) Then
                Err.Raise vbObjectError + 7560, "VisualTeX", _
                    "The image Equation paragraph contains a duplicate SEQ field."
            End If
        Next candidateField
        Exit Sub
    End If

    If formulaRange.InlineShapes.Count <> 0 Or _
       formulaRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The native Equation verification formula is ambiguous."
    End If
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If Not VTHelperParagraphOwnsNativeEquationSequence( _
       helperParagraph) Or _
       helperParagraph.Start < formulaParagraph.End Or _
       Not VTEquationCaptionBookmarkIsCollapsedInParagraph( _
           captionRange, helperParagraph) Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The native Equation SEQ helper paragraph is invalid" & _
            " [helper=" & CStr(helperParagraph.Start) & "-" & _
                CStr(helperParagraph.End) & _
            "; caption=" & CStr(captionRange.Start) & "-" & _
                CStr(captionRange.End) & "]."
    End If
    For Each candidateField In formulaRange.OMaths(1).Range.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The native Equation SEQ was absorbed into OMath."
        End If
    Next candidateField
    Set visibleNumberField = VTNativeEquationArrayReferenceField( _
        formulaRange, formulaId)
    If visibleNumberField Is Nothing Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The native Equation visible REF is missing."
    End If
    If Not VTNativeEquationNumberIsInsideMath( _
       formulaRange, visibleNumberField) Or _
       VTEquationSequenceResultText(visibleNumberField) <> expectedText Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The native Equation visible REF is invalid."
    End If
    Set expectedNumberRange = VTNativeEquationArrayNumberRange( _
        formulaRange, visibleNumberField)
    If expectedNumberRange Is Nothing Or _
       Not VTNativeEquationNumberBookmarkIsCompatible( _
           numberRange, formulaRange, visibleNumberField) Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The native Equation array number Bookmark is incomplete" & _
            " [bookmark=" & CStr(numberRange.Start) & "-" & _
                CStr(numberRange.End) & _
            "; formula=" & CStr(formulaRange.Start) & "-" & _
                CStr(formulaRange.End) & "]."
    End If
End Sub

Private Sub VTFinalizeParagraphEquationNumber( _
    ByVal documentObject As Document, _
    ByVal formulaRange As Range, _
    ByVal formulaId As String, _
    Optional ByVal deferReconcile As Boolean = False)

    Dim formulaParagraph As Range
    Dim sequenceField As Field
    Dim equationLabelName As String
    Dim sequenceBookmarkName As String
    Dim sequenceOrdinal As Long
    Dim fieldAnchor As Long

    If documentObject Is Nothing Or formulaRange Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The Equation finalization target is missing."
    End If
    Set formulaParagraph = _
        VTWordParagraphContainingFormula(formulaRange)
    If formulaParagraph Is Nothing Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "Word could not resolve the Equation finalization paragraph."
    End If
    If formulaParagraph.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The new Equation number unexpectedly entered a table."
    End If

    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    If formulaRange.InlineShapes.Count = 1 Or _
       formulaRange.OMaths.Count = 1 Then
        Set sequenceField = VTNativeEquationSequenceHelperField( _
            documentObject, formulaId)
    Else
        Set sequenceField = VTFindEquationSequenceField(formulaParagraph)
    End If
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The Equation SEQ field is missing."
    End If

    equationLabelName = VTNativeEquationLabelName()
    fieldAnchor = VTEquationFieldStart(sequenceField)
    sequenceOrdinal = VTEquationSequenceOrdinal( _
        documentObject, sequenceField, equationLabelName)
    If sequenceOrdinal < 1 Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The Equation has no document-order ordinal."
    End If
    VTApplyEquationSequenceOrdinal _
        sequenceField, equationLabelName, sequenceOrdinal
    Set sequenceField = VTResolveEquationSequenceFieldNear( _
        documentObject, fieldAnchor, 64)
    VTRefreshParagraphEquationBookmarks _
        documentObject, sequenceField, formulaId
    If deferReconcile Then
        Set formulaRange = VTNumberedFormulaRangeForId( _
            documentObject, formulaId)
        If formulaRange Is Nothing Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "Word lost the migrated image Equation before deferred reconciliation."
        End If
        VTVerifyParagraphEquationNumberIntegrity _
            formulaRange, formulaId, sequenceOrdinal
        VTEnsureOrphanWatchScheduled
        Exit Sub
    End If

    VTReconcileEquationNumbers documentObject, fieldAnchor
    Set sequenceField = VTEquationSequenceFieldForBookmark( _
        documentObject, sequenceBookmarkName)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "Word lost the Equation SEQ during reconciliation."
    End If
    sequenceOrdinal = VTEquationSequenceOrdinal( _
        documentObject, sequenceField, equationLabelName)
    If sequenceOrdinal < 1 Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The reconciled Equation SEQ has no ordinal."
    End If
    VTRefreshParagraphEquationBookmarks _
        documentObject, sequenceField, formulaId
    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "Word lost the visible Equation during reconciliation."
    End If
    VTVerifyParagraphEquationNumberIntegrity _
        formulaRange, formulaId, sequenceOrdinal
    VTEnsureOrphanWatchScheduled
End Sub

Private Sub VTDeleteEquationCaptionText( _
    ByVal documentObject As Document, _
    ByVal formulaId As String)

    Dim bookmarkName As String
    Dim captionRange As Range
    Dim helperParagraph As Range
    Dim preserveNativeHelper As Boolean

    If documentObject Is Nothing Then Exit Sub
    bookmarkName = VTEquationCaptionBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(bookmarkName) Then Exit Sub
    Set captionRange = documentObject.Bookmarks(bookmarkName).Range.Duplicate
    If captionRange.End > captionRange.Start Then
        Set helperParagraph = captionRange.Paragraphs(1).Range.Duplicate
        preserveNativeHelper = _
            VTHelperParagraphOwnsNativeEquationSequence(helperParagraph)
    End If
    documentObject.Bookmarks(bookmarkName).Delete
    If captionRange.End > captionRange.Start And _
       Not preserveNativeHelper Then captionRange.Delete
End Sub

Private Function VTFindEquationSequenceField( _
    ByVal paragraphRange As Range) As Field

    Dim candidate As Field
    Dim match As Field
    Dim matchCount As Long
    Dim equationLabelName As String
    Dim tableEnd As Long
    Dim candidateDistance As Long
    Dim bestDistance As Long

    If paragraphRange Is Nothing Then Exit Function
    equationLabelName = VTNativeEquationLabelName()
    For Each candidate In paragraphRange.Fields
        If VTIsNativeEquationSequenceField(candidate, equationLabelName) Then
            matchCount = matchCount + 1
            Set match = candidate
        End If
    Next candidate
    If matchCount = 0 And paragraphRange.Information(wdWithInTable) Then
        tableEnd = paragraphRange.Tables(1).Range.End
        bestDistance = 2147483647
        For Each candidate In paragraphRange.Document.Fields
            If VTIsNativeEquationSequenceField( _
               candidate, equationLabelName) Then
                candidateDistance = VTEquationFieldStart(candidate) - tableEnd
                If candidateDistance >= 0 And candidateDistance <= 64 Then
                    If candidateDistance < bestDistance Then
                        bestDistance = candidateDistance
                        matchCount = 1
                        Set match = candidate
                    ElseIf candidateDistance = bestDistance Then
                        matchCount = matchCount + 1
                    End If
                End If
            End If
        Next candidate
    End If
    If matchCount > 1 Then
        Err.Raise vbObjectError + 7494, "VisualTeX", _
            "The formula paragraph contains more than one Equation number."
    End If
    If matchCount = 1 Then Set VTFindEquationSequenceField = match
End Function

Private Function VTEquationLayoutWidth( _
    ByVal paragraphRange As Range) As Single

    Dim sectionObject As Section
    Dim textWidth As Single

    If paragraphRange Is Nothing Then
        Err.Raise vbObjectError + 7425, "VisualTeX", _
            "The Equation layout paragraph is missing."
    End If
    On Error Resume Next
    Set sectionObject = paragraphRange.Sections(1)
    textWidth = sectionObject.PageSetup.TextColumns.Width
    If textWidth <= 0! Then
        textWidth = sectionObject.PageSetup.PageWidth - _
            sectionObject.PageSetup.LeftMargin - _
            sectionObject.PageSetup.RightMargin
    End If
    On Error GoTo 0
    If textWidth <= 2! Then
        Err.Raise vbObjectError + 7425, "VisualTeX", _
            "Word returned an invalid text width for equation numbering."
    End If
    VTEquationLayoutWidth = textWidth
End Function

Private Sub VTConfigureNumberedEquationParagraph( _
    ByVal paragraphRange As Range)

    Dim textWidth As Single

    textWidth = VTEquationLayoutWidth(paragraphRange)
    ' Applying Caption repeatedly resets direct character formatting in Word
    ' for Mac, including the deliberate vertical Equation-number correction.
    ' Apply it only when the paragraph first becomes a numbered formula; later
    ' refreshes update paragraph geometry without touching formula/number fonts.
    If paragraphRange.Style <> wdStyleCaption Then
        paragraphRange.Style = wdStyleCaption
    End If
    With paragraphRange.ParagraphFormat
        .Alignment = wdAlignParagraphLeft
        .LeftIndent = 0!
        .RightIndent = 0!
        .FirstLineIndent = 0!
        .SpaceBefore = 0!
        .SpaceAfter = 0!
        .LineSpacingRule = wdLineSpaceSingle
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .WidowControl = True
        .TabStops.ClearAll
        .TabStops.Add _
            Position:=textWidth / 2!, _
            Alignment:=wdAlignTabCenter, _
            Leader:=wdTabLeaderSpaces
        .TabStops.Add _
            Position:=textWidth - 1!, _
            Alignment:=wdAlignTabRight, _
            Leader:=wdTabLeaderSpaces
    End With
End Sub

Private Sub VTUpdateEquationCaptionText( _
    ByVal documentObject As Document, _
    ByVal sequenceField As Field, _
    ByVal formulaId As String, _
    ByVal captionText As String)

    Dim closingRange As Range
    Dim fieldAnchor As Long
    Dim fieldEnd As Long

    If documentObject Is Nothing Or sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7495, "VisualTeX", _
            "The Equation caption target is missing."
    End If
    fieldAnchor = VTEquationFieldStart(sequenceField)
    VTDeleteEquationCaptionText documentObject, formulaId

    Set sequenceField = VTResolveEquationSequenceFieldNear( _
        documentObject, fieldAnchor, 256)
    fieldEnd = VTEquationFieldEnd(sequenceField)
    If fieldEnd >= documentObject.Content.End Or _
       documentObject.Range(fieldEnd, fieldEnd + 1).Text <> ")" Then
        Set closingRange = documentObject.Range( _
            Start:=fieldEnd, End:=fieldEnd)
        closingRange.Text = ")"
    End If
End Sub

Private Sub VTNormalizeEquationNumberLayout( _
    ByVal formulaRange As Range, _
    ByVal renderedHeightPoints As Double, _
    ByVal formulaId As String, _
    ByVal captionText As String)

    Dim sequenceField As Field

    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7496, "VisualTeX", _
            "The numbered formula Range is missing."
    End If
    Set sequenceField = VTFindEquationSequenceField( _
        formulaRange.Paragraphs(1).Range)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7498, "VisualTeX", _
            "The Equation number field is missing from the formula paragraph."
    End If
    VTNormalizeEquationNumberLayoutWithField _
        formulaRange, renderedHeightPoints, formulaId, captionText, _
        VTEquationFieldStart(sequenceField)
End Sub

Private Sub VTNormalizeEquationNumberLayoutWithField( _
    ByVal formulaRange As Range, _
    ByVal renderedHeightPoints As Double, _
    ByVal formulaId As String, _
    ByVal captionText As String, _
    ByVal fieldAnchor As Long)

    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim prefixRange As Range
    Dim separatorRange As Range
    Dim insertionRange As Range
    Dim sequenceField As Field
    Dim nativeEquation As OMath
    Dim openingRange As Range
    Dim resultRange As Range
    Dim closingRange As Range
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim centerTabStart As Long
    Dim parenStart As Long
    Dim numberEnd As Long
    Dim paragraphStart As Long
    Dim formulaStart As Long
    Dim formulaLength As Long
    Dim formulaWasImage As Boolean
    Dim rightProbeStart As Long
    Dim rightProbeText As String
    Dim separatorText As String
    Dim preferredSize As Single
    Dim numberRaisePoints As Single
    Dim operationStage As String
    Dim operationErrorNumber As Long
    Dim operationErrorDescription As String

    On Error GoTo NormalizeFailed
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7496, "VisualTeX", _
            "The numbered formula Range is missing."
    End If

    operationStage = "configure-paragraph"
    Set documentObject = formulaRange.Document
    formulaWasImage = formulaRange.InlineShapes.Count = 1
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    paragraphStart = paragraphRange.Start

    operationStage = "remove-old-caption-text"
    VTDeleteEquationCaptionText documentObject, formulaId
    If formulaWasImage Then
        Set formulaRange = VTResolveImageFormulaInParagraph( _
            documentObject, paragraphStart).Range.Duplicate
    Else
        Set formulaRange = VTResolveNativeEquationRange( _
            documentObject, paragraphStart, 512)
    End If
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    paragraphStart = paragraphRange.Start
    VTConfigureNumberedEquationParagraph paragraphRange

    operationStage = "normalize-center-prefix"
    formulaLength = formulaRange.End - formulaRange.Start
    If formulaLength <= 0 Then
        Err.Raise vbObjectError + 7496, "VisualTeX", _
            "The numbered formula Range has no content."
    End If
    Set prefixRange = documentObject.Range( _
        Start:=paragraphRange.Start, End:=formulaRange.Start)
    If VTWordRangeHasMeaningfulText(prefixRange) Then
        Err.Raise vbObjectError + 7497, "VisualTeX", _
            "A numbered display formula must occupy its own paragraph."
    End If
    If prefixRange.End > prefixRange.Start Then prefixRange.Delete
    formulaStart = paragraphStart
    If formulaWasImage Then
        Set formulaRange = VTResolveImageFormulaInParagraph( _
            documentObject, paragraphStart).Range.Duplicate
        Set formulaRange = _
            VTPrependCenterTabPreservingImage(formulaRange)
        formulaStart = formulaRange.Start
        formulaLength = formulaRange.End - formulaRange.Start
    Else
        Set formulaRange = VTResolveNativeEquationRange( _
            documentObject, formulaRange.Start, 16)
        Set formulaRange = VTPrependCenterTabPreservingNativeFormula( _
            formulaRange, formulaId)
        formulaStart = formulaRange.Start
        formulaLength = formulaRange.End - formulaRange.Start
    End If
    operationStage = "verify-formula-after-center-prefix"
    If formulaWasImage And formulaRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7539, "VisualTeX", _
            "Word deleted the formula image while inserting the center tab" & _
            " [range=" & CStr(formulaRange.Start) & "-" & _
            CStr(formulaRange.End) & "; documentImages=" & _
            CStr(documentObject.InlineShapes.Count) & "]."
    End If

    If fieldAnchor < 0 Then
        Err.Raise vbObjectError + 7498, "VisualTeX", _
            "The Equation number field anchor is invalid."
    End If

    operationStage = "right-separator-resolve-field-start"
    Set sequenceField = VTResolveEquationSequenceFieldNear( _
        documentObject, fieldAnchor, 256)
    fieldStart = VTEquationFieldStart(sequenceField)
    fieldAnchor = fieldStart

    operationStage = "right-separator-create-range"
    Set separatorRange = documentObject.Range( _
        Start:=formulaRange.End, End:=fieldStart)

    operationStage = "right-separator-inspect-text"
    separatorText = separatorRange.Text
    separatorText = Replace$(separatorText, vbTab, "")
    separatorText = Replace$(separatorText, " ", "")
    separatorText = Replace$(separatorText, ChrW(160), "")
    separatorText = Replace$(separatorText, "(", "")
    If Len(separatorText) > 0 Then
        Err.Raise vbObjectError + 7499, "VisualTeX", _
            "The Equation number separator contains unexpected text."
    End If

    operationStage = "right-separator-delete-old"
    If separatorRange.End > separatorRange.Start Then separatorRange.Delete
    operationStage = "verify-formula-after-separator-delete"
    If formulaWasImage And formulaRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7540, "VisualTeX", _
            "Word deleted the formula image while removing the old number separator."
    End If

    operationStage = "format-formula-before-final-number-boundary"
    formulaRange.Font.Position = 0
    If formulaRange.OMaths.Count = 1 Then
        ' A Word display OMath owns an m:oMathPara and therefore forces the
        ' following right-tab/number into another paragraph. Keep numbered OMML
        ' as a built-up inline OMath on its own centered display paragraph. The
        ' visible presentation remains a line formula, while formula, number and
        ' the sole paragraph mark stay in one ordinary Word paragraph.
        Set nativeEquation = formulaRange.OMaths(1)
        nativeEquation.Type = wdOMathInline
        nativeEquation.BuildUp
        nativeEquation.Range.Font.Position = 0
        nativeEquation.Range.Font.Size = _
            VTPreferredNativeDisplayFontSize(nativeEquation.Range)
        Set formulaRange = VTResolveNativeEquationRange( _
            documentObject, paragraphStart, 512)
        If formulaRange.OMaths.Count <> 1 Or _
           formulaRange.OMaths(1).Type <> wdOMathInline Then
            Err.Raise vbObjectError + 7544, "VisualTeX", _
                "Word did not retain the numbered OMML as one built-up inline formula."
        End If
        formulaStart = formulaRange.Start
        formulaLength = formulaRange.End - formulaRange.Start
        fieldAnchor = VTEquationFieldStart(sequenceField)
    End If

    operationStage = "right-separator-insert-canonical"
    If formulaWasImage Then
        Set insertionRange = documentObject.Range( _
            Start:=formulaRange.End, End:=formulaRange.End)
        insertionRange.InsertBefore vbTab & "("
        Set formulaRange = documentObject.Range( _
            Start:=formulaStart, End:=formulaStart + formulaLength)
    Else
        VTPlaceCaretAfterInlineNativeEquation formulaRange
        Selection.TypeText Text:=vbTab & "("
        Set formulaRange = VTResolveNativeEquationRange( _
            documentObject, formulaStart, 16)
        Set separatorRange = documentObject.Range( _
            Start:=formulaRange.End, End:=formulaRange.End + 2)
        If separatorRange.Text <> vbTab & "(" Or _
           separatorRange.OMaths.Count <> 0 Then
            Err.Raise vbObjectError + 7544, "VisualTeX", _
                "Word could not persist the native Equation number boundary outside OMath."
        End If
    End If
    operationStage = "verify-formula-after-separator-insert"
    If formulaWasImage And formulaRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7541, "VisualTeX", _
            "Word deleted the formula image while inserting the canonical number separator."
    End If

    operationStage = "refresh-paragraph-after-structure"
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    VTConfigureNumberedEquationParagraph paragraphRange

    operationStage = "refresh-field-before-caption"
    Set sequenceField = VTResolveEquationSequenceFieldNear( _
        documentObject, fieldAnchor, 256)
    fieldAnchor = VTEquationFieldStart(sequenceField)

    operationStage = "update-caption-text"
    VTUpdateEquationCaptionText _
        documentObject, sequenceField, formulaId, captionText
    If formulaWasImage Then
        Set formulaRange = VTResolveImageFormulaInParagraph( _
            documentObject, paragraphStart).Range.Duplicate
    Else
        Set formulaRange = VTResolveNativeEquationRange( _
            documentObject, paragraphStart, 512)
    End If

    operationStage = "resolve-visible-number-ranges"
    Set sequenceField = VTResolveEquationSequenceFieldNear( _
        documentObject, fieldAnchor, 256)
    fieldStart = VTEquationFieldStart(sequenceField)
    fieldAnchor = fieldStart
    fieldEnd = VTEquationFieldEnd(sequenceField)
    parenStart = fieldStart - 1
    numberEnd = fieldEnd + 1
    Set openingRange = documentObject.Range( _
        Start:=parenStart, End:=fieldStart)
    Set resultRange = sequenceField.Result.Duplicate
    Set closingRange = documentObject.Range( _
        Start:=fieldEnd, End:=numberEnd)

    operationStage = "format-visible-number"
    preferredSize = VTPreferredEquationFontSize(formulaRange, True)
    openingRange.Font.Size = preferredSize
    resultRange.Font.Size = preferredSize
    closingRange.Font.Size = preferredSize
    If formulaRange.OMaths.Count = 1 Then
        numberRaisePoints = VTNativeEquationNumberRaisePoints( _
            formulaRange, preferredSize)
    Else
        numberRaisePoints = VTEquationNumberRaisePoints( _
            renderedHeightPoints, preferredSize)
    End If
    openingRange.Font.Position = CLng(numberRaisePoints)
    resultRange.Font.Position = CLng(numberRaisePoints)
    closingRange.Font.Position = CLng(numberRaisePoints)

    operationStage = "bookmark-visible-number"
    VTSetEquationNumberBookmark _
        documentObject, formulaId, openingRange, closingRange

    operationStage = "verify-layout-tabs"
    Set sequenceField = VTResolveEquationSequenceFieldNear( _
        documentObject, fieldAnchor, 256)
    fieldAnchor = VTEquationFieldStart(sequenceField)
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    centerTabStart = paragraphRange.Start
    If paragraphRange.Document.Range( _
        centerTabStart, centerTabStart + 1).Text <> vbTab Then
        Err.Raise vbObjectError + 7500, "VisualTeX", _
            "Word did not persist the center-tab Equation layout."
    End If
    fieldStart = fieldAnchor
    If fieldStart < 2 Or _
       documentObject.Range( _
           fieldStart - 2, _
           fieldStart - 1).Text <> vbTab Then
        rightProbeStart = fieldStart - 4
        If rightProbeStart < 0 Then rightProbeStart = 0
        rightProbeText = documentObject.Range( _
            rightProbeStart, fieldStart).Text
        rightProbeText = Replace$(rightProbeText, vbTab, "<TAB>")
        rightProbeText = Replace$(rightProbeText, vbCr, "<CR>")
        Err.Raise vbObjectError + 7501, "VisualTeX", _
            "Word did not persist the right-tab Equation layout" & _
            " [fieldStart=" & CStr(fieldStart) & _
            "; formulaEnd=" & CStr(formulaRange.End) & _
            "; before=" & rightProbeText & "]."
    End If
    Exit Sub

NormalizeFailed:
    operationErrorNumber = Err.Number
    operationErrorDescription = Err.Description
    Err.Raise operationErrorNumber, "VisualTeX Equation numbering", _
        "VTNormalizeEquationNumberLayoutWithField/" & operationStage & _
        ": " & operationErrorDescription
End Sub

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

    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & ": native display Equation is missing."
    End If
    If formulaRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & ": native display Equation is ambiguous."
    End If
    Set documentObject = formulaRange.Document
    Set formulaRange = formulaRange.OMaths(1).Range.Duplicate
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
    If sequenceRange.Text <> sequenceField.Result.Text Or _
       VTEquationSequenceResultText(visibleNumberField) <> _
           VTEquationSequenceResultText(sequenceField) Or _
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
       Trim$(CStr(nativeItems(nativeItemIndex))) <> _
           VTEquationSequenceResultText(sequenceField) Then
        Err.Raise vbObjectError + 7565, "VisualTeX", _
            assertionName & _
            ": Word's native Equation list is not the pure number."
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
        Set formulaRange = paragraphRange.OMaths(1).Range.Duplicate
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
                    Set visibleNumberField = VTImageEquationReferenceField( _
                        formulaRange, formulaId)
                    If Not visibleNumberField Is Nothing Then
                        Set numberRange = VTImageEquationNumberRange( _
                            formulaRange, visibleNumberField)
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
        End If
    Next numberBookmark
    If Len(bookmarkName) = 0 Or visibleNumberField Is Nothing Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The numbered image Equation invariant REF is missing."
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

    fieldStart = VTEquationFieldStart(visibleNumberField)
    fieldEnd = VTEquationFieldEnd(visibleNumberField)
    Set openingRange = documentObject.Range(fieldStart - 1, fieldStart)
    Set closingRange = documentObject.Range(fieldEnd, fieldEnd + 1)
    Set formulaStartProbe = formulaRange.Duplicate
    formulaStartProbe.Collapse wdCollapseStart
    Set formulaEndProbe = formulaRange.Duplicate
    formulaEndProbe.Collapse wdCollapseEnd
    Set numberEndProbe = closingRange.Duplicate
    numberEndProbe.Collapse wdCollapseEnd
    documentObject.Repaginate

    VTNumberedEquationInvariantSnapshot = _
        "paragraph=" & CStr(paragraphRange.Start) & ":" & _
            CStr(paragraphRange.End) & _
        "|text=" & Replace$(Replace$(paragraphRange.Text, vbTab, "<TAB>"), vbCr, "<CR>") & _
        "|visibleRef=" & CStr(fieldStart) & ":" & CStr(fieldEnd) & _
            ":" & visibleNumberField.Result.Text & ":" & _
                Trim$(visibleNumberField.Code.Text) & _
        "|externalSeq=" & CStr(VTEquationFieldStart(sequenceField)) & _
            ":" & CStr(VTEquationFieldEnd(sequenceField)) & _
            ":" & sequenceField.Result.Text & ":" & _
                Trim$(sequenceField.Code.Text) & _
        "|helper=" & CStr(helperParagraph.Start) & ":" & _
            CStr(helperParagraph.End)
    VTNumberedEquationInvariantSnapshot = VTNumberedEquationInvariantSnapshot & _
        "|parens=" & openingRange.Text & closingRange.Text & _
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
        "|bookmark=" & bookmarkName & ":" & _
            documentObject.Range(openingRange.Start, closingRange.End).Text
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
            layoutTable.Cell(1, 2).Range.OMaths(1).Range.Duplicate
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
            CStr(sequenceRange.Start) & ":" & CStr(sequenceRange.End) & _
            ":" & sequenceField.Result.Text & ":" & _
            Trim$(sequenceField.Code.Text) & _
        "|paragraphs=" & _
            CStr(layoutTable.Cell(1, 2).Range.Paragraphs(1).Range.Start) & _
            ":" & _
            CStr(layoutTable.Cell(1, 2).Range.Paragraphs(1).Range.End) & _
            ":" & _
            CStr(layoutTable.Cell(1, 3).Range.Paragraphs(1).Range.Start) & _
            ":" & _
            CStr(layoutTable.Cell(1, 3).Range.Paragraphs(1).Range.End) & _
        "|tabs=" & _
            CStr(VTCustomTabStopCount( _
                layoutTable.Cell(1, 2).Range.Paragraphs(1).Range)) & _
            ":" & _
            CStr(VTCustomTabStopCount( _
                layoutTable.Cell(1, 3).Range.Paragraphs(1).Range)) & _
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

Private Function VTEnsureImageEquationNumber( _
    ByRef formulaShape As InlineShape, _
    ByVal renderedHeightPoints As Double, _
    ByVal formulaId As String, _
    ByVal captionText As String, _
    ByRef numberCreated As Boolean) As Range

    Dim formulaRange As Range
    Dim layoutTable As Table
    Dim numberBookmarkName As String

    If formulaShape Is Nothing Then
        Err.Raise vbObjectError + 7502, "VisualTeX", _
            "The numbered formula image is missing."
    End If
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    numberCreated = Not formulaShape.Range.Document.Bookmarks.Exists( _
        numberBookmarkName)
    Set formulaRange = formulaShape.Range.Duplicate

    ' Preserve existing three-cell formulas in old documents, but every new or
    ' already-migrated formula uses one ordinary Word paragraph with a centered
    ' tab for the formula and a right-aligned tab for the native SEQ number.
    If formulaRange.Information(wdWithInTable) Then
        Set layoutTable = formulaRange.Tables(1)
        If layoutTable.Rows.Count = 1 And layoutTable.Columns.Count = 3 Then
            Set formulaShape = layoutTable.Cell(1, 2).Range.InlineShapes(1)
            VTEnsureEquationNumberFields layoutTable, formulaId
            Set VTEnsureImageEquationNumber = layoutTable.Range.Duplicate
            Exit Function
        End If
    End If

    Set VTEnsureImageEquationNumber = VTInsertEquationNumber( _
        formulaShape, formulaId, captionText)
End Function

Private Function VTEnsureNativeEquationArrayNumber( _
    ByVal equationRange As Range, _
    ByVal formulaId As String) As Range

    Dim documentObject As Document
    Dim nativeEquation As OMath
    Dim exactEquationRange As Range
    Dim formulaContentRange As Range
    Dim paragraphRange As Range
    Dim prefixRange As Range
    Dim suffixRange As Range
    Dim insertionRange As Range
    Dim markerRange As Range
    Dim numberSlotRange As Range
    Dim numberRange As Range
    Dim legacyTailRange As Range
    Dim sequenceField As Field
    Dim legacySequenceField As Field
    Dim visibleNumberField As Field
    Dim candidateField As Field
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim captionBookmarkName As String
    Dim formulaStart As Long
    Dim suffixText As String
    Dim operationStage As String
    Dim operationErrorNumber As Long
    Dim operationErrorDescription As String

    On Error GoTo ArrayFailed
    If equationRange Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The native Equation array target is missing."
    End If
    If equationRange.OMaths.Count <> 1 Or _
       Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The native Equation array target is invalid."
    End If
    Set nativeEquation = equationRange.OMaths(1)
    Set documentObject = nativeEquation.Range.Document
    Set exactEquationRange = nativeEquation.Range.Duplicate
    formulaStart = exactEquationRange.Start
    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)

    operationStage = "reuse-external-sequence-array"
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    Set visibleNumberField = VTNativeEquationArrayReferenceField( _
        exactEquationRange, formulaId)
    If Not sequenceField Is Nothing Then
        If Not visibleNumberField Is Nothing Then
            If VTNativeEquationNumberIsInsideMath( _
               exactEquationRange, visibleNumberField) Then
        nativeEquation.BuildUp
        nativeEquation.Type = wdOMathDisplay
        nativeEquation.Justification = wdOMathJcCenterGroup
        Set exactEquationRange = VTResolveNativeEquationRange( _
            documentObject, formulaStart, 128)
        Set visibleNumberField = VTNativeEquationArrayReferenceField( _
            exactEquationRange, formulaId)
        If visibleNumberField Is Nothing Then
            Err.Raise vbObjectError + 7563, "VisualTeX", _
                "Word lost the native Equation number REF during BuildUp."
        End If
        Set numberRange = VTNativeEquationArrayNumberRange( _
            exactEquationRange, visibleNumberField)
        If numberRange Is Nothing Then
            Err.Raise vbObjectError + 7563, "VisualTeX", _
                "Word lost the native Equation array number range."
        End If
        VTSetEquationNumberBookmarkExact _
            documentObject, formulaId, numberRange
        VTRefreshParagraphEquationBookmarks _
            documentObject, sequenceField, formulaId
        VTSetNativeFormulaBookmark _
            documentObject, exactEquationRange, formulaId
        VTFinalizeParagraphEquationNumber _
            documentObject, exactEquationRange, formulaId
        Set exactEquationRange = VTNumberedFormulaRangeForId( _
            documentObject, formulaId)
        If exactEquationRange Is Nothing Then
            Err.Raise vbObjectError + 7563, "VisualTeX", _
                "Word lost the reused native Equation identity."
        End If
        VTSetNativeFormulaBookmark _
            documentObject, exactEquationRange, formulaId
                Set VTEnsureNativeEquationArrayNumber = _
                    VTWordParagraphContainingFormula(exactEquationRange)
                Exit Function
            End If
        End If
    End If

    operationStage = "normalize-source-formula"
    nativeEquation.Type = wdOMathInline
    nativeEquation.BuildUp
    Set exactEquationRange = nativeEquation.Range.Duplicate
    Set paragraphRange = VTWordParagraphContainingFormula(exactEquationRange)
    If paragraphRange Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word could not resolve the native Equation paragraph before numbering."
    End If

    ' r19 and earlier placed the true SEQ inside OMath. Strip the complete old
    ' array tail before constructing the external-SEQ/internal-REF architecture.
    Set markerRange = VTNativeEquationArrayMarkerRange(exactEquationRange)
    Set legacySequenceField = VTFindEquationSequenceField(paragraphRange)
    If markerRange Is Nothing And Not legacySequenceField Is Nothing Then
        operationStage = "detach-image-paragraph-sequence"
        Set sequenceField = VTEquationSequenceFieldForBookmark( _
            documentObject, sequenceBookmarkName)
        If sequenceField Is Nothing Then
            Err.Raise vbObjectError + 7563, "VisualTeX", _
                "The trailing Equation SEQ has no VisualTeX identity."
        End If
        If VTEquationFieldStart(sequenceField) <> _
           VTEquationFieldStart(legacySequenceField) Then
            Err.Raise vbObjectError + 7563, "VisualTeX", _
                "The trailing Equation SEQ does not belong to this formula."
        End If
        If documentObject.Bookmarks.Exists(numberBookmarkName) Then
            documentObject.Bookmarks(numberBookmarkName).Delete
        End If
        If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
            documentObject.Bookmarks(sequenceBookmarkName).Delete
        End If
        If documentObject.Bookmarks.Exists(captionBookmarkName) Then
            documentObject.Bookmarks(captionBookmarkName).Delete
        End If
        Set sequenceField = Nothing
    End If
    If Not markerRange Is Nothing Then
        operationStage = "remove-legacy-array-tail"
        If documentObject.Bookmarks.Exists(numberBookmarkName) Then
            documentObject.Bookmarks(numberBookmarkName).Delete
        End If
        If Not legacySequenceField Is Nothing Then
            If VTNativeEquationSequenceIsInsideMath( _
               exactEquationRange, legacySequenceField) Then
                If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
                    documentObject.Bookmarks(sequenceBookmarkName).Delete
                End If
                If documentObject.Bookmarks.Exists(captionBookmarkName) Then
                    documentObject.Bookmarks(captionBookmarkName).Delete
                End If
                Set sequenceField = Nothing
            End If
        End If
        Set legacyTailRange = documentObject.Range( _
            Start:=markerRange.Start, End:=exactEquationRange.End)
        legacyTailRange.Delete
        Set exactEquationRange = VTResolveNativeEquationRange( _
            documentObject, formulaStart, 128)
        Set markerRange = _
            VTNativeEquationArrayMarkerRange(exactEquationRange)
        If Not markerRange Is Nothing Then
            Err.Raise vbObjectError + 7563, "VisualTeX", _
                "Word did not remove the legacy Equation array tail."
        End If
        For Each candidateField In exactEquationRange.Fields
            If VTIsNativeEquationSequenceField( _
               candidateField, VTNativeEquationLabelName()) Then
                Err.Raise vbObjectError + 7563, "VisualTeX", _
                    "The legacy Equation SEQ remained inside OMath."
            End If
        Next candidateField
    End If

    operationStage = "validate-dedicated-formula-paragraph"
    Set paragraphRange = VTWordParagraphContainingFormula(exactEquationRange)
    If paragraphRange Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word lost the dedicated native Equation paragraph."
    End If
    Set prefixRange = documentObject.Range( _
        Start:=paragraphRange.Start, End:=exactEquationRange.Start)
    If VTWordRangeHasMeaningfulText(prefixRange) Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The numbered native Equation has body text before the formula."
    End If
    If prefixRange.End > prefixRange.Start Then prefixRange.Delete
    formulaStart = paragraphRange.Start
    Set exactEquationRange = VTResolveNativeEquationRange( _
        documentObject, formulaStart, 64)
    Set paragraphRange = VTWordParagraphContainingFormula(exactEquationRange)
    Set suffixRange = documentObject.Range( _
        Start:=exactEquationRange.End, End:=paragraphRange.End - 1)
    If suffixRange.InlineShapes.Count <> 0 Or _
       suffixRange.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The numbered native Equation has another object after the formula."
    End If
    suffixText = suffixRange.Text
    For Each candidateField In suffixRange.Fields
        If Not VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            Err.Raise vbObjectError + 7563, "VisualTeX", _
                "The numbered native Equation has an unrelated trailing field."
        End If
        suffixText = Replace$( _
            suffixText, candidateField.Result.Text, "", 1, -1, _
            vbBinaryCompare)
    Next candidateField
    suffixText = Replace$(suffixText, vbTab, "")
    suffixText = Replace$(suffixText, " ", "")
    suffixText = Replace$(suffixText, ChrW(160), "")
    suffixText = Replace$(suffixText, ChrW(8203), "")
    suffixText = Replace$(suffixText, ChrW(8288), "")
    suffixText = Replace$(suffixText, "(", "")
    suffixText = Replace$(suffixText, ")", "")
    If Len(suffixText) <> 0 Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The numbered native Equation has trailing body text."
    End If
    If suffixRange.End > suffixRange.Start Then suffixRange.Delete

    operationStage = "create-external-sequence-helper"
    Set exactEquationRange = VTResolveNativeEquationRange( _
        documentObject, formulaStart, 64)
    Set sequenceField = VTEnsureNativeEquationSequenceHelper( _
        exactEquationRange, formulaId)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word did not create the native Equation SEQ helper."
    End If
    If sequenceField.Result.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word did not isolate the native Equation SEQ outside OMath."
    End If

    operationStage = "insert-equation-array-marker"
    Set exactEquationRange = VTResolveNativeEquationRange( _
        documentObject, formulaStart, 64)
    Set insertionRange = documentObject.Range( _
        Start:=exactEquationRange.End, End:=exactEquationRange.End)
    insertionRange.Select
    Selection.TypeText Text:="#()"
    Set exactEquationRange = VTResolveNativeEquationRange( _
        documentObject, formulaStart, 64)
    Set markerRange = VTNativeEquationArrayMarkerRange(exactEquationRange)
    If markerRange Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word did not keep the Equation array marker inside OMath."
    End If
    Set numberSlotRange = documentObject.Range( _
        Start:=markerRange.End + 1, End:=markerRange.End + 1)
    If numberSlotRange.Start >= exactEquationRange.End Or _
       documentObject.Range( _
           markerRange.End, markerRange.End + 1).Text <> "(" Or _
       documentObject.Range( _
           numberSlotRange.Start, numberSlotRange.Start + 1).Text <> ")" Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word did not preserve an internal Equation array number slot."
    End If

    operationStage = "insert-equation-array-reference"
    Set visibleNumberField = documentObject.Fields.Add( _
        Range:=numberSlotRange, Type:=wdFieldRef, _
        Text:=VTParenthesizedEquationReferenceFieldText( _
            sequenceBookmarkName), _
        PreserveFormatting:=False)
    visibleNumberField.Update
    Set exactEquationRange = VTResolveNativeEquationRange( _
        documentObject, formulaStart, 128)
    If Not VTNativeEquationNumberIsInsideMath( _
       exactEquationRange, visibleNumberField) Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word did not keep the Equation number REF inside OMath."
    End If
    For Each candidateField In exactEquationRange.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            Err.Raise vbObjectError + 7563, "VisualTeX", _
                "Word absorbed the external Equation SEQ into OMath."
        End If
    Next candidateField

    operationStage = "build-equation-array"
    Set nativeEquation = exactEquationRange.OMaths(1)
    nativeEquation.BuildUp
    nativeEquation.Type = wdOMathDisplay
    nativeEquation.Justification = wdOMathJcCenterGroup
    Set exactEquationRange = VTResolveNativeEquationRange( _
        documentObject, formulaStart, 128)
    Set visibleNumberField = VTNativeEquationArrayReferenceField( _
        exactEquationRange, formulaId)
    If visibleNumberField Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word lost the Equation number REF while building display math."
    End If
    If Not VTNativeEquationNumberIsInsideMath( _
       exactEquationRange, visibleNumberField) Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The Equation number REF left its OMath array during BuildUp."
    End If
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The native Equation SEQ helper disappeared after BuildUp."
    End If
    If sequenceField.Result.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The native Equation SEQ left its external helper paragraph."
    End If
    If VTEquationSequenceResultText(visibleNumberField) <> _
       VTEquationSequenceResultText(sequenceField) Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The native Equation number REF does not match its external SEQ."
    End If

    operationStage = "finalize-equation-array-identity"
    Set paragraphRange = VTWordParagraphContainingFormula(exactEquationRange)
    If paragraphRange Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word lost the native Equation paragraph after BuildUp."
    End If
    VTConfigureNativeEquationArrayParagraph paragraphRange
    Set numberRange = VTNativeEquationArrayNumberRange( _
        exactEquationRange, visibleNumberField)
    If numberRange Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word did not expose the final native Equation number range."
    End If
    VTSetEquationNumberBookmarkExact _
        documentObject, formulaId, numberRange
    VTRefreshParagraphEquationBookmarks _
        documentObject, sequenceField, formulaId
    Set exactEquationRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If exactEquationRange Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word lost the native Equation identity before finalization."
    End If
    Set formulaContentRange = _
        VTNativeEquationFormulaContentRange(exactEquationRange)
    If formulaContentRange Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word did not preserve the Equation content before its array marker."
    End If
    VTSetNativeFormulaBookmark _
        documentObject, exactEquationRange, formulaId
    VTFinalizeParagraphEquationNumber _
        documentObject, exactEquationRange, formulaId
    Set exactEquationRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If exactEquationRange Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "Word lost the finalized native Equation identity."
    End If
    VTSetNativeFormulaBookmark _
        documentObject, exactEquationRange, formulaId
    Set VTEnsureNativeEquationArrayNumber = _
        VTWordParagraphContainingFormula(exactEquationRange)
    Exit Function

ArrayFailed:
    operationErrorNumber = Err.Number
    operationErrorDescription = Err.Description
    Err.Raise operationErrorNumber, "VisualTeX native Equation array", _
        "VTEnsureNativeEquationArrayNumber/" & operationStage & _
        ": " & operationErrorDescription
End Function

Private Function VTEnsureNativeEquationNumber( _
    ByVal equationRange As Range, _
    ByVal renderedHeightPoints As Double, _
    ByVal formulaId As String, _
    ByVal captionText As String, _
    ByRef numberCreated As Boolean) As Range

    Dim formulaRange As Range
    Dim layoutTable As Table
    Dim numberBookmarkName As String

    If equationRange Is Nothing Then
        Err.Raise vbObjectError + 7470, "VisualTeX", _
            "The native equation number target is missing."
    End If
    If equationRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7470, "VisualTeX", _
            "The native equation number target is ambiguous."
    End If
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    numberCreated = Not equationRange.Document.Bookmarks.Exists( _
        numberBookmarkName)
    Set formulaRange = equationRange.OMaths(1).Range.Duplicate

    If formulaRange.Information(wdWithInTable) Then
        Set layoutTable = formulaRange.Tables(1)
        If layoutTable.Rows.Count = 1 And layoutTable.Columns.Count = 3 Then
            VTEnsureEquationNumberFields layoutTable, formulaId
            Set VTEnsureNativeEquationNumber = layoutTable.Range.Duplicate
            Exit Function
        End If
    End If

    Set VTEnsureNativeEquationNumber = _
        VTEnsureNativeEquationArrayNumber(formulaRange, formulaId)
End Function

Private Function VTEnsureNumberedDisplayTable( _
    ByVal formulaRange As Range, _
    ByVal formulaIsImage As Boolean, _
    ByVal formulaId As String) As Table

    Dim documentObject As Document
    Dim layoutTable As Table
    Dim paragraphRange As Range
    Dim beforeRange As Range
    Dim afterRange As Range
    Dim cleanupRange As Range
    Dim insertionRange As Range
    Dim centerRange As Range
    Dim backupDocument As Document
    Dim backupRange As Range
    Dim nativeEquation As OMath
    Dim paragraphStart As Long
    Dim cellXml As String
    Dim operationStage As String
    Dim operationErrorNumber As Long
    Dim operationErrorDescription As String

    On Error GoTo LayoutFailed
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The numbered display formula Range is missing."
    End If
    Set documentObject = formulaRange.Document

    If formulaRange.Information(wdWithInTable) Then
        operationStage = "reuse-table"
        Set layoutTable = formulaRange.Tables(1)
        If layoutTable.Rows.Count <> 1 Or layoutTable.Columns.Count <> 3 Then
            Err.Raise vbObjectError + 7549, "VisualTeX", _
                "The numbered display formula is inside an incompatible table."
        End If
    ElseIf Not formulaIsImage Then
        operationStage = "wrap-native-display-paragraph"
        Set layoutTable = VTWrapNativeDisplayParagraphInTable(formulaRange)
    Else
        operationStage = "validate-standalone-image-paragraph"
        Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
        Set beforeRange = paragraphRange.Duplicate
        beforeRange.End = formulaRange.Start
        Set afterRange = paragraphRange.Duplicate
        afterRange.Start = formulaRange.End
        If paragraphRange.InlineShapes.Count <> 1 Or _
           paragraphRange.OMaths.Count <> 0 Or _
           beforeRange.InlineShapes.Count <> 0 Or _
           beforeRange.OMaths.Count <> 0 Or _
           afterRange.InlineShapes.Count <> 0 Or _
           afterRange.OMaths.Count <> 0 Or _
           VTWordRangeHasMeaningfulText(beforeRange) Or _
           VTWordRangeHasMeaningfulText(afterRange) Then
            Err.Raise vbObjectError + 7549, "VisualTeX", _
                "A numbered display image must occupy its own paragraph;" & _
                " surrounding Word content was preserved."
        End If

        operationStage = "backup-formula-image"
        paragraphStart = paragraphRange.Start
        Set backupDocument = Documents.Add(Visible:=False)
        Set backupRange = backupDocument.Content.Duplicate
        backupRange.Collapse wdCollapseStart
        backupRange.FormattedText = formulaRange.FormattedText
        If backupDocument.InlineShapes.Count <> 1 Then
            Err.Raise vbObjectError + 7549, "VisualTeX", _
                "Word could not back up the numbered formula image."
        End If
        Set backupRange = backupDocument.InlineShapes(1).Range.Duplicate
        documentObject.Activate

        operationStage = "remove-legacy-bookmarks"
        If documentObject.Bookmarks.Exists( _
           VTEquationNumberBookmarkName(formulaId)) Then
            documentObject.Bookmarks( _
                VTEquationNumberBookmarkName(formulaId)).Delete
        End If
        If documentObject.Bookmarks.Exists( _
           VTEquationCaptionBookmarkName(formulaId)) Then
            documentObject.Bookmarks( _
                VTEquationCaptionBookmarkName(formulaId)).Delete
        End If

        operationStage = "clear-source-paragraph"
        Set cleanupRange = paragraphRange.Duplicate
        If cleanupRange.End > cleanupRange.Start Then
            cleanupRange.End = cleanupRange.End - 1
        End If
        If cleanupRange.End > cleanupRange.Start Then cleanupRange.Delete

        operationStage = "create-table"
        Set insertionRange = documentObject.Range( _
            Start:=paragraphStart, End:=paragraphStart)
        Set layoutTable = documentObject.Tables.Add( _
            Range:=insertionRange, NumRows:=1, NumColumns:=3)

        operationStage = "restore-formula-in-center-cell"
        Set centerRange = layoutTable.Cell(1, 2).Range.Duplicate
        centerRange.End = centerRange.End - 1
        centerRange.Collapse wdCollapseStart
        centerRange.FormattedText = backupRange.FormattedText
        backupDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set backupDocument = Nothing
    End If

    operationStage = "configure-table"
    VTConfigureNumberedDisplayTable layoutTable

    operationStage = "finalize-formula"
    If formulaIsImage Then
        If layoutTable.Cell(1, 2).Range.InlineShapes.Count <> 1 Then
            Err.Raise vbObjectError + 7549, "VisualTeX", _
                "The numbered table does not contain exactly one formula image."
        End If
        layoutTable.Cell(1, 2).Range.InlineShapes(1).Range.Font.Position = 0
    Else
        If layoutTable.Cell(1, 2).Range.OMaths.Count <> 1 Then
            Err.Raise vbObjectError + 7549, "VisualTeX", _
                "The numbered table does not contain exactly one native formula."
        End If
        Set nativeEquation = layoutTable.Cell(1, 2).Range.OMaths(1)
        nativeEquation.Type = wdOMathDisplay
        nativeEquation.Justification = wdOMathJcCenter
        nativeEquation.BuildUp
        nativeEquation.Range.Font.Position = 0
        nativeEquation.Range.Font.Size = _
            VTPreferredNativeDisplayFontSize(layoutTable.Cell(1, 2).Range)

        cellXml = layoutTable.Cell(1, 2).Range.WordOpenXML
        If layoutTable.Cell(1, 2).Range.Paragraphs.Count <> 2 Or _
           InStr(1, cellXml, "<m:oMathPara", vbBinaryCompare) = 0 Then
            operationStage = "repair-native-display-cell"
            Set layoutTable = VTRebuildExistingNativeDisplayCell(layoutTable)
        Else
            operationStage = "compact-native-display-tail"
            VTCompactNativeDisplayCellTail layoutTable
        End If
    End If

    Set VTEnsureNumberedDisplayTable = layoutTable
    Exit Function

LayoutFailed:
    operationErrorNumber = Err.Number
    operationErrorDescription = Err.Description
    On Error Resume Next
    If Not backupDocument Is Nothing Then
        backupDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise operationErrorNumber, "VisualTeX Equation table layout", _
        "VTEnsureNumberedDisplayTable/" & operationStage & ": " & _
        operationErrorDescription
End Function

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

Private Function VTRebuildExistingNativeDisplayCell( _
    ByVal layoutTable As Table) As Table

    Const temporaryBookmarkName As String = "VT_TMP_DISPLAY_REPAIR"

    Dim documentObject As Document
    Dim sourceEquation As OMath
    Dim centerEquation As OMath
    Dim insertionRange As Range
    Dim sourceParagraph As Range
    Dim centerRange As Range
    Dim trailingParagraph As Range
    Dim cleanupRange As Range
    Dim temporaryStart As Long
    Dim cellXml As String
    Dim operationStage As String
    Dim operationErrorNumber As Long
    Dim operationErrorDescription As String

    On Error GoTo RepairFailed
    operationStage = "validate-table"
    If layoutTable Is Nothing Or layoutTable.Rows.Count <> 1 Or _
       layoutTable.Columns.Count <> 3 Or _
       layoutTable.Cell(1, 2).Range.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The existing numbered native display cell is invalid."
    End If

    Set documentObject = layoutTable.Range.Document
    Set sourceEquation = layoutTable.Cell(1, 2).Range.OMaths(1)
    If documentObject.Bookmarks.Exists(temporaryBookmarkName) Then
        documentObject.Bookmarks(temporaryBookmarkName).Delete
    End If

    operationStage = "append-temporary-source"
    documentObject.Content.InsertAfter vbCr
    temporaryStart = documentObject.Content.End - 1
    Set insertionRange = documentObject.Range( _
        Start:=temporaryStart, End:=temporaryStart)
    insertionRange.FormattedText = sourceEquation.Range.FormattedText
    Set sourceParagraph = documentObject.Range( _
        Start:=temporaryStart, End:=temporaryStart).Paragraphs(1).Range.Duplicate
    If sourceParagraph.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word could not create the temporary source OMath."
    End If

    operationStage = "promote-temporary-source"
    Set sourceEquation = sourceParagraph.OMaths(1)
    sourceEquation.Type = wdOMathDisplay
    sourceEquation.Justification = wdOMathJcCenter
    sourceEquation.Range.Font.Position = 0
    sourceEquation.Range.Font.Size = _
        VTPreferredNativeDisplayFontSize(sourceEquation.Range)
    sourceEquation.BuildUp
    Set sourceParagraph = documentObject.Range( _
        Start:=temporaryStart, End:=temporaryStart).Paragraphs(1).Range.Duplicate
    If sourceParagraph.OMaths.Count <> 1 Or _
       sourceParagraph.OMaths(1).Type <> wdOMathDisplay Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word did not preserve the temporary display OMath."
    End If
    documentObject.Bookmarks.Add _
        Name:=temporaryBookmarkName, Range:=sourceParagraph

    operationStage = "transfer-display-paragraph"
    ' Replace the complete existing cell content in one operation. Deleting it
    ' first lets Word merge the imported paragraph mark into the end-of-cell
    ' marker, which leaves image-to-native conversion with only one paragraph.
    Set centerRange = layoutTable.Cell(1, 2).Range.Duplicate
    centerRange.End = centerRange.End - 1
    centerRange.FormattedText = _
        documentObject.Bookmarks(temporaryBookmarkName).Range.FormattedText

    If layoutTable.Cell(1, 2).Range.Paragraphs.Count = 1 Then
        operationStage = "append-required-tail"
        Set centerRange = layoutTable.Cell(1, 2).Range.Duplicate
        centerRange.End = centerRange.End - 1
        centerRange.Collapse wdCollapseEnd
        centerRange.InsertBefore vbCr
    End If

    operationStage = "verify-transferred-cell"
    If layoutTable.Cell(1, 2).Range.Paragraphs.Count <> 2 Or _
       layoutTable.Cell(1, 2).Range.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word did not rebuild the two-paragraph display cell" & _
            " [paragraphs=" & _
            CStr(layoutTable.Cell(1, 2).Range.Paragraphs.Count) & _
            "; maths=" & _
            CStr(layoutTable.Cell(1, 2).Range.OMaths.Count) & "]."
    End If
    Set centerEquation = layoutTable.Cell(1, 2).Range.OMaths(1)
    If centerEquation.Type <> wdOMathDisplay Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The rebuilt center-cell OMath is not display math."
    End If
    Set trailingParagraph = _
        layoutTable.Cell(1, 2).Range.Paragraphs(2).Range.Duplicate
    If trailingParagraph.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(trailingParagraph) Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The rebuilt display-cell tail contains unexpected content."
    End If

    operationStage = "delete-temporary-source"
    Set cleanupRange = _
        documentObject.Bookmarks(temporaryBookmarkName).Range.Duplicate
    documentObject.Bookmarks(temporaryBookmarkName).Delete
    cleanupRange.Delete

    operationStage = "finalize-rebuilt-cell"
    VTConfigureNumberedDisplayTable layoutTable
    Set centerEquation = layoutTable.Cell(1, 2).Range.OMaths(1)
    centerEquation.Type = wdOMathDisplay
    centerEquation.Justification = wdOMathJcCenter
    centerEquation.BuildUp
    Set centerEquation = layoutTable.Cell(1, 2).Range.OMaths(1)
    centerEquation.Range.Font.Position = 0
    centerEquation.Range.Font.Size = _
        VTPreferredNativeDisplayFontSize(layoutTable.Cell(1, 2).Range)
    VTCompactNativeDisplayCellTail layoutTable

    operationStage = "verify-rebuilt-openxml"
    cellXml = layoutTable.Cell(1, 2).Range.WordOpenXML
    If InStr(1, cellXml, "<m:oMathPara", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The rebuilt display cell has no m:oMathPara structure."
    End If

    Set VTRebuildExistingNativeDisplayCell = layoutTable
    Exit Function

RepairFailed:
    operationErrorNumber = Err.Number
    operationErrorDescription = Err.Description
    On Error Resume Next
    If Not documentObject Is Nothing Then
        If documentObject.Bookmarks.Exists(temporaryBookmarkName) Then
            Set cleanupRange = _
                documentObject.Bookmarks(temporaryBookmarkName).Range.Duplicate
            documentObject.Bookmarks(temporaryBookmarkName).Delete
            cleanupRange.Delete
        End If
    End If
    On Error GoTo 0
    Err.Raise operationErrorNumber, _
        "VisualTeX native display repair", _
        "VTRebuildExistingNativeDisplayCell/" & operationStage & ": " & _
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

Private Sub VTConfigureNumberedDisplayTable(ByVal layoutTable As Table)
    Dim cellIndex As Long

    If layoutTable Is Nothing Or layoutTable.Rows.Count <> 1 Or _
       layoutTable.Columns.Count <> 3 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The numbered display table must be one row by three columns."
    End If
    With layoutTable
        .AllowAutoFit = False
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100!
        .Borders.Enable = False
        .Rows.AllowBreakAcrossPages = False
    End With
    For cellIndex = 1 To 3
        layoutTable.Columns(cellIndex).PreferredWidthType = _
            wdPreferredWidthPercent
        If cellIndex = 2 Then
            layoutTable.Columns(cellIndex).PreferredWidth = 60!
        Else
            layoutTable.Columns(cellIndex).PreferredWidth = 20!
        End If
        layoutTable.Cell(1, cellIndex).VerticalAlignment = _
            wdCellAlignVerticalCenter
        With layoutTable.Cell(1, cellIndex).Range.ParagraphFormat
            .LeftIndent = 0!
            .RightIndent = 0!
            .FirstLineIndent = 0!
            .SpaceBefore = 0!
            .SpaceAfter = 0!
            .TabStops.ClearAll
        End With
    Next cellIndex
    layoutTable.Cell(1, 1).Range.ParagraphFormat.Alignment = _
        wdAlignParagraphLeft
    layoutTable.Cell(1, 2).Range.ParagraphFormat.Alignment = _
        wdAlignParagraphCenter
    layoutTable.Cell(1, 3).Range.ParagraphFormat.Alignment = _
        wdAlignParagraphRight
End Sub

Private Sub VTEnsureEquationNumberFields( _
    ByVal layoutTable As Table, _
    ByVal formulaId As String)

    Dim documentObject As Document
    Dim sequenceField As Field
    Dim candidateField As Field
    Dim referenceField As Field
    Dim hiddenParagraph As Range
    Dim hiddenContent As Range
    Dim rightContent As Range
    Dim fieldRange As Range
    Dim numberRange As Range
    Dim sequenceBookmarkName As String
    Dim captionBookmarkName As String
    Dim numberBookmarkName As String
    Dim equationLabelName As String
    Dim expectedText As String
    Dim sequenceOrdinal As Long
    Dim visibleStart As Long
    Dim helperAnchor As Long

    If layoutTable Is Nothing Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The Equation number table is missing."
    End If
    Set documentObject = layoutTable.Range.Document
    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)

    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        For Each candidateField In documentObject.Fields
            If VTIsNativeEquationSequenceField( _
               candidateField, VTNativeEquationLabelName()) Then
                If candidateField.Result.Start <= _
                   documentObject.Bookmarks( _
                       sequenceBookmarkName).Range.Start And _
                   candidateField.Result.End >= _
                   documentObject.Bookmarks( _
                       sequenceBookmarkName).Range.End Then
                    Set sequenceField = candidateField
                    Exit For
                End If
            End If
        Next candidateField
    End If

    If sequenceField Is Nothing Then
        If documentObject.Bookmarks.Exists(captionBookmarkName) Then
            documentObject.Bookmarks(captionBookmarkName).Delete
        End If
        If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
            documentObject.Bookmarks(sequenceBookmarkName).Delete
        End If
        Set hiddenParagraph = _
            VTInsertDedicatedEquationHelperParagraph(layoutTable)
        Set fieldRange = documentObject.Range( _
            Start:=hiddenParagraph.Start, End:=hiddenParagraph.Start)
        equationLabelName = VTNativeEquationLabelName()
        Set sequenceField = VTInsertRegisteredEquationCaption( _
            fieldRange, equationLabelName)
        sequenceOrdinal = VTEquationSequenceOrdinal( _
            documentObject, sequenceField, equationLabelName)
        If sequenceOrdinal < 1 Then
            Err.Raise vbObjectError + 7549, "VisualTeX", _
                "Word registered the native Equation SEQ field outside the" & _
                " document sequence."
        End If
        VTApplyEquationSequenceOrdinal _
            sequenceField, equationLabelName, sequenceOrdinal
        documentObject.Bookmarks.Add _
            name:=sequenceBookmarkName, _
            Range:=sequenceField.Result.Duplicate
        Set hiddenParagraph = sequenceField.Result.Paragraphs(1).Range.Duplicate
        documentObject.Bookmarks.Add _
            name:=captionBookmarkName, Range:=hiddenParagraph
    Else
        equationLabelName = VTNativeEquationLabelName()
        sequenceOrdinal = VTEquationSequenceOrdinal( _
            documentObject, sequenceField, equationLabelName)
        If sequenceOrdinal < 1 Then
            Err.Raise vbObjectError + 7549, "VisualTeX", _
                "The existing native Equation SEQ field has no ordinal."
        End If
        VTApplyEquationSequenceOrdinal _
            sequenceField, equationLabelName, sequenceOrdinal
        If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
            documentObject.Bookmarks(sequenceBookmarkName).Delete
        End If
        documentObject.Bookmarks.Add _
            name:=sequenceBookmarkName, _
            Range:=sequenceField.Result.Duplicate
        Set hiddenParagraph = sequenceField.Result.Paragraphs(1).Range.Duplicate
    End If

    helperAnchor = hiddenParagraph.Start
    VTFormatHiddenEquationParagraph hiddenParagraph

    ' Reconcile this formula and every following native Equation before
    ' creating the new visible right-cell REF. The incremental pass performs
    ' the SEQ update and mirror refresh once for each affected formula.
    ' Updating SEQ fields can invalidate a Bookmark that
    ' wraps a field result on Word for Mac, so immediately re-resolve the known
    ' helper paragraph and restore the exact VT_N_/VT_C_ pair afterward.
    VTReconcileEquationNumbers documentObject, helperAnchor
    Set sequenceField = VTResolveEquationSequenceFieldNear( _
        documentObject, helperAnchor, 128)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word lost the native Equation SEQ field during reconciliation."
    End If
    equationLabelName = VTNativeEquationLabelName()
    sequenceOrdinal = VTEquationSequenceOrdinal( _
        documentObject, sequenceField, equationLabelName)
    If sequenceOrdinal < 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The reconciled native Equation SEQ field has no ordinal."
    End If
    VTApplyEquationSequenceOrdinal _
        sequenceField, equationLabelName, sequenceOrdinal
    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        documentObject.Bookmarks(sequenceBookmarkName).Delete
    End If
    documentObject.Bookmarks.Add _
        name:=sequenceBookmarkName, Range:=sequenceField.Result.Duplicate
    Set hiddenParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If documentObject.Bookmarks.Exists(captionBookmarkName) Then
        documentObject.Bookmarks(captionBookmarkName).Delete
    End If
    documentObject.Bookmarks.Add _
        name:=captionBookmarkName, Range:=hiddenParagraph
    VTFormatHiddenEquationParagraph hiddenParagraph

    ' The right cell uses the same validated Windows structure: ordinary
    ' parentheses around a native REF to the native SEQ result Bookmark. VT_R_
    ' bookmarks only this visible (n) range for formula identity and geometry;
    ' it is no longer a second numbering source or a cross-reference mirror.
    Set rightContent = layoutTable.Cell(1, 3).Range.Duplicate
    rightContent.End = rightContent.End - 1
    If rightContent.End > rightContent.Start Then rightContent.Delete
    visibleStart = layoutTable.Cell(1, 3).Range.Start
    Set rightContent = documentObject.Range( _
        Start:=visibleStart, End:=visibleStart)
    rightContent.Text = "()"
    Set fieldRange = documentObject.Range( _
        Start:=visibleStart + 1, End:=visibleStart + 1)
    Set referenceField = documentObject.Fields.Add( _
        Range:=fieldRange, Type:=wdFieldRef, _
        Text:=VTParenthesizedEquationReferenceFieldText( _
            sequenceBookmarkName), _
        PreserveFormatting:=False)
    referenceField.Update
    expectedText = CStr(sequenceOrdinal)
    If VTEquationSequenceResultText(referenceField) <> expectedText Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word did not format the visible Equation REF from the native" & _
            " SEQ result" & _
            " [code=" & referenceField.Code.Text & _
            "; result=" & referenceField.Result.Text & _
            "; expected=" & expectedText & "]."
    End If
    Set numberRange = documentObject.Range( _
        Start:=visibleStart, _
        End:=VTEquationFieldEnd(referenceField) + 1)
    If numberRange.Text <> "(" & expectedText & ")" Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word did not preserve the complete visible Equation number" & _
            " [text=" & numberRange.Text & _
            "; expected=(" & expectedText & ")]."
    End If
    VTSetEquationNumberBookmarkExact _
        documentObject, formulaId, numberRange
    VTFormatVisibleEquationReference _
        documentObject, referenceField, numberRange
    layoutTable.Cell(1, 3).Range.ParagraphFormat.Alignment = _
        wdAlignParagraphRight

    ' Native SEQ/REF insertion and field reconciliation can let Word for Mac
    ' re-evaluate the table layout. Reapply the final invariant only after all
    ' field mutations are complete so the public result remains 100% wide,
    ' fixed 20/60/20, borderless, and non-breaking across pages.
    VTConfigureNumberedDisplayTable layoutTable
    VTVerifyEquationNumberFieldIntegrity _
        layoutTable, formulaId, sequenceOrdinal
    VTEnsureOrphanWatchScheduled
End Sub

Private Function VTVisibleEquationNumberFontSize( _
    ByVal documentObject As Document) As Single

    Dim normalSize As Single

    If documentObject Is Nothing Then
        VTVisibleEquationNumberFontSize = 11!
        Exit Function
    End If
    On Error Resume Next
    normalSize = documentObject.Styles(wdStyleNormal).Font.Size
    On Error GoTo 0
    If normalSize <= 0! Or normalSize > 72! Then normalSize = 11!
    VTVisibleEquationNumberFontSize = normalSize
End Function

Private Sub VTFormatBodyEquationReference( _
    ByVal documentObject As Document, _
    ByVal referenceField As Field, _
    ByVal referenceRange As Range)

    Dim visibleSize As Single

    If documentObject Is Nothing Or referenceField Is Nothing Or _
       referenceRange Is Nothing Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The body Equation reference formatting target is missing."
    End If
    visibleSize = VTVisibleEquationNumberFontSize(documentObject)
    With referenceField.Result.Font
        .Hidden = False
        .Color = wdColorAutomatic
        .Position = 0
        .Size = visibleSize
    End With
    With referenceRange.Font
        .Hidden = False
        .Color = wdColorAutomatic
        .Position = 0
        .Size = visibleSize
    End With
    ' Word's native InsertCrossReference may add MERGEFORMAT. That switch is
    ' valid for body references; visibility is enforced from the actual result.
    If referenceField.Result.Font.Hidden <> False Or _
       referenceField.Result.Font.Color <> wdColorAutomatic Or _
       Abs(referenceField.Result.Font.Size - visibleSize) > 0.1 Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The body Equation reference inherited invalid formatting."
    End If
End Sub

Private Sub VTNormalizeBodyEquationReferenceVisibility( _
    ByVal documentObject As Document, _
    Optional ByVal changedFrom As Long = -1, _
    Optional ByVal updateFields As Boolean = False)

    If documentObject Is Nothing Then Exit Sub
    VTNormalizeBodyEquationReferenceVisibilityInRange _
        documentObject, documentObject.Content, changedFrom, updateFields
End Sub

Private Sub VTNormalizeBodyEquationReferenceVisibilityInRange( _
    ByVal documentObject As Document, _
    ByVal scanRange As Range, _
    Optional ByVal changedFrom As Long = -1, _
    Optional ByVal updateFields As Boolean = False)

    Dim candidateField As Field
    Dim targetBookmarkName As String
    Dim sequenceBookmarkName As String
    Dim formulaId As String
    Dim targetKind As String
    Dim formulaIds As Variant
    Dim formulaIdsLoaded As Boolean
    Dim targetStart As Long
    Dim isVisualTeXReference As Boolean
    Dim shouldUpdateField As Boolean
    Dim shouldFormatVisualTeX As Boolean

    If documentObject Is Nothing Or scanRange Is Nothing Then Exit Sub
    For Each candidateField In scanRange.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            targetBookmarkName = VTReferenceTargetBookmarkName( _
                candidateField.Code.Text)
            formulaId = VTFormulaIdFromSequenceBookmarkName( _
                targetBookmarkName)
            isVisualTeXReference = (Len(formulaId) > 0)
            If Not isVisualTeXReference And _
               Left$(targetBookmarkName, 4) = "_Ref" Then
                If Not formulaIdsLoaded Then
                    formulaIds = VTValidNumberedFormulaIds(documentObject)
                    formulaIdsLoaded = True
                End If
                formulaId = VTFormulaIdForReferenceTarget( _
                    documentObject, targetBookmarkName, targetKind, _
                    candidateField.Result.Text, formulaIds)
                isVisualTeXReference = (Len(formulaId) > 0)
            End If

            targetStart = -1
            If isVisualTeXReference Then
                sequenceBookmarkName = _
                    VTEquationSequenceNumberBookmarkName(formulaId)
                If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
                    targetStart = documentObject.Bookmarks( _
                        sequenceBookmarkName).Range.Start
                End If
            ElseIf documentObject.Bookmarks.Exists(targetBookmarkName) Then
                ' Preserve Word-native Equation cross-references as well as
                ' VisualTeX references. Their private target Bookmark still
                ' provides a stable affected-position boundary.
                targetStart = documentObject.Bookmarks( _
                    targetBookmarkName).Range.Start
            End If

            shouldUpdateField = updateFields
            If shouldUpdateField And changedFrom >= 0 Then
                shouldUpdateField = _
                    (targetStart >= changedFrom)
            End If
            If shouldUpdateField Then candidateField.Update

            shouldFormatVisualTeX = isVisualTeXReference
            If shouldFormatVisualTeX And changedFrom >= 0 Then
                shouldFormatVisualTeX = _
                    (targetStart >= changedFrom)
            End If
            If shouldFormatVisualTeX Then
                VTFormatBodyEquationReference _
                    documentObject, candidateField, _
                    candidateField.Result.Duplicate
            End If
        End If
    Next candidateField
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

Private Sub VTFormatVisibleEquationReference( _
    ByVal documentObject As Document, _
    ByVal referenceField As Field, _
    ByVal numberRange As Range)

    Dim visibleSize As Single

    If documentObject Is Nothing Or referenceField Is Nothing Or _
       numberRange Is Nothing Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The visible Equation number formatting target is missing."
    End If
    visibleSize = VTVisibleEquationNumberFontSize(documentObject)
    With referenceField.Result.Font
        .Hidden = False
        .Color = wdColorAutomatic
        .Position = 0
        .Size = visibleSize
    End With
    With numberRange.Font
        .Hidden = False
        .Color = wdColorAutomatic
        .Position = 0
        .Size = visibleSize
    End With
    With numberRange.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBefore = 0!
        .SpaceAfter = 0!
    End With
    If InStr(1, referenceField.Code.Text, "MERGEFORMAT", _
       vbTextCompare) > 0 Or _
       referenceField.Result.Font.Hidden <> False Or _
       referenceField.Result.Font.Color <> wdColorAutomatic Or _
       Abs(referenceField.Result.Font.Size - visibleSize) > 0.1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The visible Equation number inherited hidden helper formatting" & _
            " [code=" & referenceField.Code.Text & _
            "; size=" & CStr(referenceField.Result.Font.Size) & _
            "; expectedSize=" & CStr(visibleSize) & "]."
    End If
End Sub

Private Sub VTVerifyEquationNumberFieldIntegrity( _
    ByVal layoutTable As Table, _
    ByVal formulaId As String, _
    ByVal expectedOrdinal As Long)

    Dim documentObject As Document
    Dim sequenceField As Field
    Dim referenceField As Field
    Dim candidateField As Field
    Dim numberRange As Range
    Dim sequenceBookmarkName As String
    Dim captionBookmarkName As String
    Dim numberBookmarkName As String
    Dim expectedNumber As String
    Dim visibleSize As Single

    If layoutTable Is Nothing Or expectedOrdinal < 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The Equation number integrity target is invalid."
    End If
    Set documentObject = layoutTable.Range.Document
    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    expectedNumber = CStr(expectedOrdinal)

    If Not documentObject.Bookmarks.Exists(sequenceBookmarkName) Or _
       Not documentObject.Bookmarks.Exists(captionBookmarkName) Or _
       Not documentObject.Bookmarks.Exists(numberBookmarkName) Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word did not preserve the complete VT_N_/VT_C_/VT_R_" & _
            " Equation Bookmark set."
    End If
    Set sequenceField = VTEquationSequenceFieldForBookmark( _
        documentObject, sequenceBookmarkName)
    If sequenceField Is Nothing Or _
       VTEquationSequenceResultText(sequenceField) <> expectedNumber Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The native Equation SEQ Bookmark has no usable number result."
    End If

    For Each candidateField In layoutTable.Cell(1, 3).Range.Fields
        If candidateField.Type = wdFieldRef And _
           InStr(1, candidateField.Code.Text, sequenceBookmarkName, _
           vbTextCompare) > 0 Then
            Set referenceField = candidateField
            Exit For
        End If
    Next candidateField
    If referenceField Is Nothing Or _
       VTEquationSequenceResultText(referenceField) <> expectedNumber Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The visible Equation REF has no usable native number result."
    End If
    Set numberRange = documentObject.Bookmarks( _
        numberBookmarkName).Range.Duplicate
    If numberRange.Text <> "(" & expectedNumber & ")" Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The visible Equation number is incomplete" & _
            " [text=" & numberRange.Text & _
            "; expected=(" & expectedNumber & ")]."
    End If
    visibleSize = VTVisibleEquationNumberFontSize(documentObject)
    If InStr(1, referenceField.Code.Text, "MERGEFORMAT", _
       vbTextCompare) > 0 Or _
       referenceField.Result.Font.Hidden <> False Or _
       referenceField.Result.Font.Color <> wdColorAutomatic Or _
       Abs(referenceField.Result.Font.Size - visibleSize) > 0.1 Or _
       numberRange.Font.Hidden <> False Or _
       numberRange.Font.Color <> wdColorAutomatic Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The Equation number exists but is not visibly formatted" & _
            " [code=" & referenceField.Code.Text & _
            "; size=" & CStr(referenceField.Result.Font.Size) & _
            "; expectedSize=" & CStr(visibleSize) & "]."
    End If
End Sub

Private Function VTBookmarkNameInFieldCode( _
    ByVal documentObject As Document, _
    ByVal fieldCode As String, _
    ByVal bookmarkPrefix As String) As String

    Dim candidateBookmark As Bookmark

    If documentObject Is Nothing Or Len(bookmarkPrefix) = 0 Then Exit Function
    For Each candidateBookmark In documentObject.Bookmarks
        If Left$(candidateBookmark.Name, Len(bookmarkPrefix)) = _
           bookmarkPrefix Then
            If InStr(1, fieldCode, candidateBookmark.Name, _
               vbTextCompare) > 0 Then
                VTBookmarkNameInFieldCode = candidateBookmark.Name
                Exit Function
            End If
        End If
    Next candidateBookmark
End Function

Private Sub VTNormalizeVisibleEquationReferenceField( _
    ByVal documentObject As Document, _
    ByVal referenceField As Field, _
    ByVal sequenceBookmarkName As String)

    Dim cellRange As Range
    Dim scaffoldRange As Range
    Dim fieldRange As Range
    Dim numberRange As Range
    Dim expectedNumber As String
    Dim expectedText As String
    Dim actualCellText As String
    Dim fieldCode As String
    Dim numberBookmarkName As String
    Dim suffixText As String
    Dim cellStart As Long
    Dim requiresCellRebuild As Boolean

    If documentObject Is Nothing Or referenceField Is Nothing Or _
       Len(sequenceBookmarkName) = 0 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The visible Equation REF normalization target is missing."
    End If
    If Not documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The native Equation number Bookmark is missing."
    End If
    If Not referenceField.Result.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "A visible Equation REF field is outside a table."
    End If
    If referenceField.Result.Tables.Count <> 1 Or _
       referenceField.Result.Tables(1).Columns.Count <> 3 Or _
       referenceField.Result.Cells(1).ColumnIndex <> 3 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "A visible Equation REF field is outside the right numbered cell."
    End If

    suffixText = Mid$(sequenceBookmarkName, _
        Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX) + 1)
    numberBookmarkName = VT_WORD_NUMBER_BOOKMARK_PREFIX & suffixText
    expectedNumber = Trim$( _
        documentObject.Bookmarks(sequenceBookmarkName).Range.Text)
    expectedText = "(" & expectedNumber & ")"
    fieldCode = referenceField.Code.Text
    Set cellRange = referenceField.Result.Cells(1).Range.Duplicate
    If cellRange.End > cellRange.Start Then cellRange.End = cellRange.End - 1
    actualCellText = cellRange.Text
    actualCellText = Replace$(actualCellText, vbCr, "")
    actualCellText = Replace$(actualCellText, Chr$(7), "")
    actualCellText = Trim$(actualCellText)
    requiresCellRebuild = _
        InStr(1, fieldCode, sequenceBookmarkName, vbTextCompare) = 0 Or _
        InStr(1, fieldCode, "\#", vbBinaryCompare) > 0 Or _
        InStr(1, fieldCode, "MERGEFORMAT", vbTextCompare) > 0 Or _
        actualCellText <> expectedText

    If requiresCellRebuild Then
        cellStart = cellRange.Start
        If cellRange.End > cellRange.Start Then cellRange.Delete
        Set scaffoldRange = documentObject.Range( _
            Start:=cellStart, End:=cellStart)
        scaffoldRange.Text = "()"
        Set fieldRange = documentObject.Range( _
            Start:=cellStart + 1, End:=cellStart + 1)
        Set referenceField = documentObject.Fields.Add( _
            Range:=fieldRange, Type:=wdFieldRef, _
            Text:=VTParenthesizedEquationReferenceFieldText( _
                sequenceBookmarkName), _
            PreserveFormatting:=False)
    End If
    referenceField.Update
    Set numberRange = documentObject.Range( _
        Start:=referenceField.Result.Cells(1).Range.Start, _
        End:=VTEquationFieldEnd(referenceField) + 1)

    If VTEquationSequenceResultText(referenceField) <> expectedNumber Or _
       numberRange.Text <> expectedText Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The visible Equation REF field did not retain the native number" & _
            " [code=" & referenceField.Code.Text & _
            "; result=" & referenceField.Result.Text & _
            "; text=" & numberRange.Text & _
            "; expected=" & expectedText & "]."
    End If
    If documentObject.Bookmarks.Exists(numberBookmarkName) Then
        documentObject.Bookmarks(numberBookmarkName).Delete
    End If
    documentObject.Bookmarks.Add _
        name:=numberBookmarkName, Range:=numberRange
    VTFormatVisibleEquationReference _
        documentObject, referenceField, numberRange
    referenceField.Result.Cells(1).Range.ParagraphFormat.Alignment = _
        wdAlignParagraphRight
End Sub

Private Sub VTReconcileEquationNumbers( _
    ByVal documentObject As Document, _
    Optional ByVal changedFrom As Long = -1)

    Dim candidate As Field
    Dim candidateBookmark As Bookmark
    Dim equationLabelName As String
    Dim fieldCode As String
    Dim sequenceBookmarkName As String
    Dim sequenceBookmarkNames() As String
    Dim sequenceAnchors() As Long
    Dim sequenceCount As Long
    Dim sequenceOrdinal As Long
    Dim referenceBookmarkNames() As String
    Dim referenceAnchors() As Long
    Dim referenceCount As Long
    Dim itemIndex As Long
    Dim shouldUpdate As Boolean
    Dim bookmarkParagraph As Range
    Dim nativeFormulaRange As Range
    Dim nativeNumberField As Field
    Dim formulaId As String
    Dim nativeBookmarkCompatible As Boolean

    If documentObject Is Nothing Then Exit Sub
    equationLabelName = VTNativeEquationLabelName()

    ' Phase 1a: capture stable anchors and VisualTeX identities without changing
    ' the Fields collection. Word for Mac can invalidate a For Each enumerator
    ' when SEQ/REF fields are updated or rebuilt during the same traversal,
    ' causing skipped numbers or runtime error 5941.
    For Each candidate In documentObject.Fields
        If VTIsNativeEquationSequenceField(candidate, equationLabelName) Then
            sequenceCount = sequenceCount + 1
            If sequenceCount = 1 Then
                ReDim sequenceAnchors(1 To 1)
                ReDim sequenceBookmarkNames(1 To 1)
            Else
                ReDim Preserve sequenceAnchors(1 To sequenceCount)
                ReDim Preserve sequenceBookmarkNames(1 To sequenceCount)
            End If
            sequenceAnchors(sequenceCount) = _
                VTEquationFieldStart(candidate)
            sequenceBookmarkNames(sequenceCount) = _
                VTSequenceBookmarkNameForField(documentObject, candidate)
        End If
    Next candidate

    ' Phase 1b: re-resolve one field at a time from its durable VT_N_ Bookmark or
    ' captured local anchor. No live Fields enumerator survives a field update.
    For itemIndex = 1 To sequenceCount
        sequenceOrdinal = itemIndex
        sequenceBookmarkName = sequenceBookmarkNames(itemIndex)
        Set candidate = Nothing
        If Len(sequenceBookmarkName) > 0 Then
            Set candidate = VTEquationSequenceFieldForBookmark( _
                documentObject, sequenceBookmarkName)
        Else
            Set candidate = VTResolveEquationSequenceFieldNear( _
                documentObject, sequenceAnchors(itemIndex), 128)
        End If
        If candidate Is Nothing Then
            Err.Raise vbObjectError + 7549, "VisualTeX", _
                "The Equation sequence field disappeared during reconciliation."
        End If

        shouldUpdate = _
            (changedFrom < 0 Or _
             VTEquationFieldStart(candidate) >= changedFrom)
        If shouldUpdate Then
            If Len(sequenceBookmarkName) > 0 Then
                ' VTRefreshEquationNumberMirror performs the SEQ update and
                ' immediately restores VT_N_/VT_C_; do not update it twice.
                VTRefreshEquationNumberMirror _
                    documentObject, candidate, sequenceBookmarkName, _
                    sequenceOrdinal
            Else
                ' Ordinary Word Equation captions remain native content. They
                ' participate in the shared sequence but are never rewritten.
                candidate.Update
            End If
        End If
    Next itemIndex

    ' Phase 2a: snapshot the legacy table-based visible REF fields. The
    ' normalizer may delete and recreate a REF, so no live Fields enumerator may
    ' remain active while normalization runs.
    For Each candidate In documentObject.Fields
        If candidate.Type = wdFieldRef And _
           candidate.Result.Information(wdWithInTable) Then
            fieldCode = candidate.Code.Text
            sequenceBookmarkName = VTReferenceTargetBookmarkName(fieldCode)
            If Left$(sequenceBookmarkName, _
               Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)) = _
               VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX And _
               documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
                referenceCount = referenceCount + 1
                If referenceCount = 1 Then
                    ReDim referenceAnchors(1 To 1)
                    ReDim referenceBookmarkNames(1 To 1)
                Else
                    ReDim Preserve referenceAnchors(1 To referenceCount)
                    ReDim Preserve referenceBookmarkNames(1 To referenceCount)
                End If
                referenceAnchors(referenceCount) = _
                    VTEquationFieldStart(candidate)
                referenceBookmarkNames(referenceCount) = _
                    sequenceBookmarkName
            End If
        End If
    Next candidate

    ' Phase 2b: resolve and normalize one visible REF at a time.
    For itemIndex = 1 To referenceCount
        sequenceBookmarkName = referenceBookmarkNames(itemIndex)
        shouldUpdate = (changedFrom < 0)
        If Not shouldUpdate And _
           documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
            shouldUpdate = documentObject.Bookmarks( _
                sequenceBookmarkName).Range.Start >= changedFrom
        End If
        If shouldUpdate Then
            Set candidate = VTResolveVisibleEquationReferenceFieldNear( _
                documentObject, referenceAnchors(itemIndex), _
                sequenceBookmarkName, 256)
            VTNormalizeVisibleEquationReferenceField _
                documentObject, candidate, sequenceBookmarkName
        End If
    Next itemIndex

    ' Phase 3: refresh body REF fields only when their target lies at or after
    ' the changed Equation position. This preserves both VisualTeX and Word-
    ' native Equation cross-references without updating unrelated earlier REF.
    VTNormalizeBodyEquationReferenceVisibility _
        documentObject, changedFrom, True

    ' Validate the durable visible-number identity once after all mutations.
    For Each candidateBookmark In documentObject.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_NUMBER_BOOKMARK_PREFIX Then
            nativeBookmarkCompatible = False
            If Left$(candidateBookmark.Range.Text, 1) <> "(" Or _
               Right$(candidateBookmark.Range.Text, 1) <> ")" Then
                Set bookmarkParagraph = _
                    candidateBookmark.Range.Paragraphs(1).Range.Duplicate
                Set nativeFormulaRange = Nothing
                Set nativeNumberField = Nothing
                formulaId = VTFormulaIdFromBookmarkSuffix( _
                    Mid$(candidateBookmark.Name, _
                        Len(VT_WORD_NUMBER_BOOKMARK_PREFIX) + 1))
                If bookmarkParagraph.OMaths.Count = 1 And _
                   Len(formulaId) > 0 Then
                    Set nativeFormulaRange = _
                        bookmarkParagraph.OMaths(1).Range.Duplicate
                    Set nativeNumberField = _
                        VTNativeEquationArrayReferenceField( _
                            nativeFormulaRange, formulaId)
                End If
                If Not nativeFormulaRange Is Nothing And _
                   Not nativeNumberField Is Nothing Then
                    nativeBookmarkCompatible = _
                        VTNativeEquationNumberBookmarkIsCompatible( _
                            candidateBookmark.Range.Duplicate, _
                            nativeFormulaRange, nativeNumberField)
                End If
                If Not nativeBookmarkCompatible Then
                    Err.Raise vbObjectError + 7549, "VisualTeX", _
                        "A visible Equation number Bookmark changed during" & _
                        " reconciliation" & _
                        " [bookmark=" & candidateBookmark.Name & _
                        "; text=" & candidateBookmark.Range.Text & _
                        "; range=" & CStr(candidateBookmark.Range.Start) & _
                        "-" & CStr(candidateBookmark.Range.End) & "]."
                End If
            End If
        End If
    Next candidateBookmark
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

Private Sub VTNormalizeUnnumberedDisplayParagraph( _
    ByVal formulaRange As Range)

    Dim paragraphRange As Range

    If formulaRange Is Nothing Then Exit Sub
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    paragraphRange.Style = wdStyleNormal
    With paragraphRange.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .LeftIndent = 0!
        .RightIndent = 0!
        .FirstLineIndent = 0!
        .TabStops.ClearAll
    End With
End Sub

Private Sub VTDeleteTrailingInlineNativeSeparator( _
    ByVal equationRange As Range)

    Dim paragraphRange As Range
    Dim trailingRange As Range
    Dim characterRange As Range
    Dim contentEnd As Long
    Dim characterValue As String

    If equationRange Is Nothing Then Exit Sub
    Set paragraphRange = equationRange.Paragraphs(1).Range.Duplicate
    contentEnd = paragraphRange.End
    Do While contentEnd > paragraphRange.Start
        characterValue = paragraphRange.Document.Range( _
            Start:=contentEnd - 1, End:=contentEnd).Text
        If characterValue = vbCr Or characterValue = Chr$(7) Then
            contentEnd = contentEnd - 1
        Else
            Exit Do
        End If
    Loop
    If equationRange.End >= contentEnd Then Exit Sub

    Set trailingRange = paragraphRange.Document.Range( _
        Start:=equationRange.End, End:=contentEnd)
    If VTWordRangeHasMeaningfulText(trailingRange) Then Exit Sub

    Do While equationRange.End < contentEnd
        Set characterRange = paragraphRange.Document.Range( _
            Start:=equationRange.End, End:=equationRange.End + 1)
        characterValue = characterRange.Text
        If characterValue = " " Or characterValue = vbTab Or _
           characterValue = ChrW(160) Or characterValue = ChrW(8203) Or _
           characterValue = ChrW(8288) Then
            characterRange.Delete
            contentEnd = contentEnd - 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub VTPlaceCaretAfterInlineNativeEquation( _
    ByVal equationRange As Range)

    Dim documentObject As Document
    Dim exactEquationRange As Range
    Dim paragraphRange As Range
    Dim caretRange As Range
    Dim anchorRange As Range

    If equationRange Is Nothing Then Exit Sub
    VTDeleteTrailingInlineNativeSeparator equationRange
    If equationRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7530, "VisualTeX", _
            "The inline equation caret target is missing."
    End If

    Set exactEquationRange = equationRange.OMaths(1).Range.Duplicate
    Set documentObject = exactEquationRange.Document
    Set paragraphRange = VTWordParagraphContainingFormula(exactEquationRange)
    If paragraphRange Is Nothing Then
        Err.Raise vbObjectError + 7530, "VisualTeX", _
            "Word could not resolve the inline equation paragraph."
    End If
    Set caretRange = documentObject.Range( _
        Start:=exactEquationRange.End, End:=exactEquationRange.End)
    caretRange.Font.Position = 0
    caretRange.Select

    ' Word for Mac keeps a collapsed Range at OMath.Range.End inside the math
    ' zone even when ordinary text already exists before the formula. Two pure
    ' Range transactions were verified by the real host and were both absorbed
    ' into OMath. MoveRight is therefore the required host operation for leaving
    ' the equation; the temporary anchor is then strictly verified outside OMath
    ' and selected so the user's first typed character replaces it.
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
    Selection.TypeText Text:=ChrW(8288)
    If Selection.Start <= paragraphRange.Start Then
        Err.Raise vbObjectError + 7530, "VisualTeX", _
            "Word did not move beyond the inline equation."
    End If
    Set anchorRange = documentObject.Range( _
        Start:=Selection.Start - 1, End:=Selection.Start)
    If anchorRange.Text <> ChrW(8288) Or _
       anchorRange.OMaths.Count <> 0 Or _
       anchorRange.Paragraphs(1).Range.Start <> paragraphRange.Start Then
        anchorRange.Delete
        Err.Raise vbObjectError + 7530, "VisualTeX", _
            "Word could not establish a text boundary after the inline equation."
    End If
    anchorRange.Font.Position = 0
    anchorRange.Select
End Sub

Private Function VTPreferredEquationFontSize( _
    ByVal contextRange As Range, _
    ByVal displaySizing As Boolean) As Single

    Dim contextSize As Single
    Dim normalSize As Single
    Dim preferredSize As Single

    On Error Resume Next
    contextSize = contextRange.Font.Size
    normalSize = ActiveDocument.Styles(wdStyleNormal).Font.Size
    On Error GoTo 0

    If normalSize <= 0! Or normalSize > 72! Then normalSize = 12!
    If contextSize <= 0! Or contextSize > 72! Then contextSize = normalSize

    If displaySizing Then
        preferredSize = normalSize
        If preferredSize < 12! Then preferredSize = 12!
    Else
        preferredSize = contextSize
        If preferredSize < 10.5 Then preferredSize = 10.5
    End If
    If preferredSize > 18! Then preferredSize = 18!
    VTPreferredEquationFontSize = preferredSize
End Function

Private Function VTPreferredNativeDisplayFontSize( _
    ByVal contextRange As Range) As Single

    Dim normalSize As Single

    On Error Resume Next
    normalSize = contextRange.Document.Styles(wdStyleNormal).Font.Size
    On Error GoTo 0

    If normalSize <= 0! Or normalSize > 72! Then normalSize = 12!
    ' Keep the nominal document text size. Display math should differ through
    ' Word's mathematical display style, not by scaling every glyph equally.
    VTPreferredNativeDisplayFontSize = normalSize
End Function

Private Function VTInsertNativeEquationAtTarget( _
    ByVal target As InlineShape, _
    ByVal ommlBase64 As String, _
    ByVal nativeDocumentPath As String, _
    ByVal displayMode As String, _
    ByVal displaySizing As Boolean, _
    ByVal replaceTarget As Boolean) As Range

    If target Is Nothing Then
        Err.Raise vbObjectError + 7450, "VisualTeX", "The native-equation insertion target is missing."
    End If
    Set VTInsertNativeEquationAtTarget = VTInsertNativeEquationAtRange( _
        target.Range, _
        ommlBase64, _
        nativeDocumentPath, _
        displayMode, _
        displaySizing, _
        replaceTarget)
End Function

Private Function VTInsertNativeEquationAtRange( _
    ByVal targetRange As Range, _
    ByVal ommlBase64 As String, _
    ByVal nativeDocumentPath As String, _
    ByVal displayMode As String, _
    ByVal displaySizing As Boolean, _
    ByVal replaceTarget As Boolean) As Range

    Dim ommlXml As String
    Dim insertionRange As Range
    Dim equationRange As Range
    Dim candidateRange As Range
    Dim targetDocument As Document
    Dim stagingDocument As Document
    Dim stagingEquationRange As Range
    Dim replacementBackupDocument As Document
    Dim replacementBackupRange As Range
    Dim replacementRange As Range
    Dim nativeEquation As OMath
    Dim failedReplacementMath As OMath
    Dim candidateMath As OMath
    Dim existingMaths As Collection
    Dim insertionStart As Long
    Dim preferredSize As Single
    Dim matchCount As Long
    Dim beforeMathCount As Long
    Dim stagingEquationLength As Long
    Dim candidateDistance As Long
    Dim bestDistance As Long
    Dim replacementApplied As Boolean
    Dim conversionErrorNumber As Long
    Dim conversionErrorDescription As String

    If targetRange Is Nothing Then
        Err.Raise vbObjectError + 7450, "VisualTeX", "The native-equation insertion target is missing."
    End If
    Set targetDocument = targetRange.Document
    If displayMode <> "inline" And displayMode <> "block" Then
        Err.Raise vbObjectError + 7451, "VisualTeX", "The native-equation display mode is invalid."
    End If
    If Not VTIsBase64UrlPayload(ommlBase64) Then
        Err.Raise vbObjectError + 7433, "VisualTeX", "The stored Word OMML payload is invalid."
    End If

    ommlXml = VTBase64UrlDecodeUtf8(ommlBase64)
    VTValidateOmmlFragment ommlXml
    If Len(nativeDocumentPath) = 0 Then
        Err.Raise vbObjectError + 7434, "VisualTeX", "This formula has no native Word staging document. Edit and save it once, then try again."
    End If
    VTValidateAbsoluteVisualTeXPath nativeDocumentPath
    If Not VTPathFileExists(nativeDocumentPath) Then
        Err.Raise vbObjectError + 7434, "VisualTeX", "The native Word staging document is missing. Edit and save the formula again."
    End If
    If displaySizing Or displayMode = "block" Then
        preferredSize = VTPreferredNativeDisplayFontSize(targetRange)
    Else
        preferredSize = VTPreferredEquationFontSize(targetRange, False)
    End If
    insertionStart = targetRange.Start
    Set existingMaths = New Collection
    For Each candidateMath In targetDocument.OMaths
        existingMaths.Add candidateMath
    Next candidateMath
    beforeMathCount = existingMaths.Count
    Set insertionRange = targetRange.Duplicate
    If Not replaceTarget Then insertionRange.Collapse wdCollapseStart

    On Error GoTo RollbackConversion

    If replaceTarget Then
        Set replacementBackupDocument = Documents.Add(Visible:=False)
        Set replacementBackupRange = replacementBackupDocument.Content
        replacementBackupRange.Collapse wdCollapseStart
        replacementBackupRange.FormattedText = targetRange.FormattedText
        If replacementBackupDocument.OMaths.Count = 1 Then
            Set replacementBackupRange = _
                replacementBackupDocument.OMaths(1).Range.Duplicate
        ElseIf replacementBackupDocument.InlineShapes.Count = 1 Then
            Set replacementBackupRange = _
                replacementBackupDocument.InlineShapes(1).Range.Duplicate
        Else
            Err.Raise vbObjectError + 7472, "VisualTeX", _
                "Word could not back up the native-equation replacement target."
        End If
        targetDocument.Activate
    End If

    ' Word for Mac 16.89.1 rejects both a bare m:oMath node and a wrapped
' WordprocessingML document through Word's XML insertion API with runtime error
    ' 6145. The companion therefore creates a real minimal DOCX ZIP. Open
    ' that package and transfer Word's parsed OMath with FormattedText. This
    ' preserves structural fractions, n-ary operators, radicals, matrices and
    ' scripts without touching the clipboard or falling back to UnicodeMath.
    Set stagingDocument = Documents.Open( _
        FileName:=nativeDocumentPath, _
        ConfirmConversions:=False, _
        ReadOnly:=True, _
        AddToRecentFiles:=False, _
        Visible:=False)
    If stagingDocument.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7434, "VisualTeX", "Word did not parse exactly one native equation from the OMML payload."
    End If
    Set stagingEquationRange = stagingDocument.OMaths(1).Range.Duplicate
    stagingEquationLength = stagingEquationRange.End - stagingEquationRange.Start
    insertionRange.FormattedText = stagingEquationRange.FormattedText
    replacementApplied = replaceTarget
    stagingDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set stagingDocument = Nothing

    ' FormattedText does not reliably expand the caller's Range on every Word
    ' for Mac build. Complex n-ary equations can also begin a few characters
    ' away from the nominal insertion boundary. Resolve the closest new OMath
    ' inside the transferred payload span instead of requiring Start <= +1.
    If replaceTarget Then
        If targetDocument.OMaths.Count < beforeMathCount Then
            Err.Raise vbObjectError + 7434, "VisualTeX", _
                "Word did not replace the native equation from the OMML payload."
        End If
    ElseIf targetDocument.OMaths.Count < beforeMathCount + 1 Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "Word did not add a native equation from the OMML payload."
    End If
    bestDistance = 2147483647
    For Each candidateMath In targetDocument.OMaths
        If replaceTarget Or _
           Not VTOMathCollectionContains(existingMaths, candidateMath) Then
            Set candidateRange = candidateMath.Range.Duplicate
            candidateDistance = Abs(candidateRange.Start - insertionStart)
            If candidateRange.Start >= insertionStart - 2 And _
               candidateRange.Start <= insertionStart + stagingEquationLength + 4 Then
                If candidateDistance < bestDistance Then
                    bestDistance = candidateDistance
                    matchCount = 1
                    Set nativeEquation = candidateMath
                ElseIf candidateDistance = bestDistance Then
                    matchCount = matchCount + 1
                End If
            End If
        End If
    Next candidateMath
    If matchCount <> 1 Or nativeEquation Is Nothing Then
        Err.Raise vbObjectError + 7434, "VisualTeX", "Word did not insert exactly one native equation from the OMML payload."
    End If

    Set equationRange = nativeEquation.Range.Duplicate
    equationRange.Font.Position = 0
    equationRange.Font.Size = preferredSize
    If displayMode = "inline" Then
        Set equationRange = _
            VTFinalizeInlineNativeEquation(nativeEquation.Range.Duplicate)
    Else
        nativeEquation.Type = wdOMathDisplay
        nativeEquation.Justification = wdOMathJcCenter
        VTNormalizeUnnumberedDisplayParagraph nativeEquation.Range.Duplicate
    End If

    Set VTInsertNativeEquationAtRange = nativeEquation.Range.Duplicate
    If Not replacementBackupDocument Is Nothing Then
        replacementBackupDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set replacementBackupDocument = Nothing
    End If
    Exit Function

RollbackConversion:
    conversionErrorNumber = Err.Number
    conversionErrorDescription = Err.Description
    On Error Resume Next
    If Not stagingDocument Is Nothing Then stagingDocument.Close SaveChanges:=wdDoNotSaveChanges
    If replaceTarget And replacementApplied And _
       Not replacementBackupRange Is Nothing Then
        If Not nativeEquation Is Nothing Then
            Set replacementRange = nativeEquation.Range.Duplicate
        Else
            Set failedReplacementMath = VTNativeMathNearStart( _
                targetDocument, insertionStart, stagingEquationLength + 8)
            If Not failedReplacementMath Is Nothing Then
                Set replacementRange = failedReplacementMath.Range.Duplicate
            End If
        End If
        If Not replacementRange Is Nothing Then
            replacementRange.FormattedText = replacementBackupRange.FormattedText
        End If
    ElseIf Not nativeEquation Is Nothing Then
        nativeEquation.Range.Delete
    End If
    If Not replacementBackupDocument Is Nothing Then
        replacementBackupDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise conversionErrorNumber, "VisualTeX Word native equation insertion", conversionErrorDescription
End Function

Private Function VTPromoteNativeEquationToDisplay( _
    ByVal equationRange As Range) As Range

    Dim nativeEquation As OMath
    Dim exactRange As Range

    If equationRange Is Nothing Or equationRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7473, "VisualTeX", _
            "VisualTeX cannot promote a missing native equation to display mode."
    End If
    Set nativeEquation = equationRange.OMaths(1)
    nativeEquation.Type = wdOMathDisplay
    nativeEquation.Justification = wdOMathJcCenter
    Set exactRange = nativeEquation.Range.Duplicate
    exactRange.Font.Position = 0
    VTNormalizeUnnumberedDisplayParagraph exactRange
    Set VTPromoteNativeEquationToDisplay = nativeEquation.Range.Duplicate
End Function

Private Function VTFinalizeInlineNativeEquation( _
    ByVal equationRange As Range) As Range

    Dim nativeEquation As OMath
    Dim exactRange As Range

    If equationRange Is Nothing Or equationRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7475, "VisualTeX", _
            "VisualTeX cannot finalize a missing inline native equation."
    End If
    Set nativeEquation = equationRange.OMaths(1)
    nativeEquation.Type = wdOMathInline
    Set exactRange = nativeEquation.Range.Duplicate
    exactRange.Font.Position = 0
    VTNormalizeInlineNativeParagraphAlignment exactRange
    VTDeleteTrailingInlineNativeSeparator exactRange
    Set VTFinalizeInlineNativeEquation = nativeEquation.Range.Duplicate
End Function

Private Sub VTNormalizeInlineNativeParagraphAlignment( _
    ByVal equationRange As Range)

    Dim paragraphRange As Range
    Dim beforeRange As Range
    Dim afterRange As Range
    Dim contentEnd As Long
    Dim trailingText As String

    If equationRange Is Nothing Then Exit Sub
    Set paragraphRange = equationRange.Paragraphs(1).Range.Duplicate
    contentEnd = paragraphRange.End
    Do While contentEnd > paragraphRange.Start
        trailingText = paragraphRange.Document.Range( _
            Start:=contentEnd - 1, End:=contentEnd).Text
        If trailingText = vbCr Or trailingText = Chr$(7) Then
            contentEnd = contentEnd - 1
        Else
            Exit Do
        End If
    Loop
    If equationRange.Start < paragraphRange.Start Or _
       equationRange.End > contentEnd Then Exit Sub

    Set beforeRange = paragraphRange.Document.Range( _
        Start:=paragraphRange.Start, End:=equationRange.Start)
    Set afterRange = paragraphRange.Document.Range( _
        Start:=equationRange.End, End:=contentEnd)
    If Not VTWordRangeHasMeaningfulText(beforeRange) And _
       Not VTWordRangeHasMeaningfulText(afterRange) Then
        paragraphRange.ParagraphFormat.Alignment = wdAlignParagraphLeft
    End If
End Sub

Private Function VTWordRangeHasMeaningfulText( _
    ByVal candidateRange As Range) As Boolean

    Dim value As String

    If candidateRange Is Nothing Then Exit Function
    value = candidateRange.Text
    value = Replace$(value, Chr$(1), "")
    value = Replace$(value, vbTab, "")
    value = Replace$(value, vbCr, "")
    value = Replace$(value, vbLf, "")
    value = Replace$(value, Chr$(7), "")
    value = Replace$(value, ChrW(160), "")
    value = Replace$(value, ChrW(8203), "")
    value = Replace$(value, ChrW(8288), "")
    value = Replace$(value, " ", "")
    VTWordRangeHasMeaningfulText = (Len(value) > 0)
End Function

Private Function VTWordParagraphContainingFormula( _
    ByVal formulaRange As Range) As Range

    Dim exactRange As Range
    Dim probeRange As Range
    Dim paragraphRange As Range
    Dim probeStart As Long

    If formulaRange Is Nothing Then Exit Function
    If formulaRange.OMaths.Count = 1 Then
        Set exactRange = formulaRange.OMaths(1).Range.Duplicate
    Else
        Set exactRange = formulaRange.Duplicate
    End If
    If exactRange.End <= exactRange.Start Then Exit Function

    ' A collapsed Range at OMath.Start can belong to the preceding paragraph on
    ' Word for Mac. Probe one character inside the formula so ordinary body text
    ' immediately above the insertion point can never be mistaken for the
    ' formula paragraph.
    probeStart = exactRange.Start + 1
    If probeStart >= exactRange.End Then probeStart = exactRange.Start
    Set probeRange = exactRange.Document.Range( _
        Start:=probeStart, End:=probeStart)
    Set paragraphRange = probeRange.Paragraphs(1).Range.Duplicate

    If exactRange.Start < paragraphRange.Start Or _
       exactRange.End > paragraphRange.End Then
        probeStart = exactRange.End - 1
        Set probeRange = exactRange.Document.Range( _
            Start:=probeStart, End:=probeStart)
        Set paragraphRange = probeRange.Paragraphs(1).Range.Duplicate
    End If
    If exactRange.Start < paragraphRange.Start Or _
       exactRange.End > paragraphRange.End Then Exit Function

    Set VTWordParagraphContainingFormula = paragraphRange
End Function

Private Function VTNativeMathNearStart( _
    ByVal documentObject As Document, _
    ByVal expectedStart As Long, _
    ByVal maximumDistance As Long) As OMath

    Dim candidateMath As OMath
    Dim match As OMath
    Dim candidateDistance As Long
    Dim bestDistance As Long
    Dim matchCount As Long

    If documentObject Is Nothing Or maximumDistance < 0 Then Exit Function
    bestDistance = 2147483647
    For Each candidateMath In documentObject.OMaths
        candidateDistance = Abs(candidateMath.Range.Start - expectedStart)
        If candidateDistance <= maximumDistance Then
            If candidateDistance < bestDistance Then
                bestDistance = candidateDistance
                matchCount = 1
                Set match = candidateMath
            ElseIf candidateDistance = bestDistance Then
                matchCount = matchCount + 1
            End If
        End If
    Next candidateMath
    If matchCount = 1 Then Set VTNativeMathNearStart = match
End Function

Private Function VTResolveNativeEquationRange( _
    ByVal documentObject As Document, _
    ByVal expectedStart As Long, _
    ByVal maximumDistance As Long) As Range

    Dim nativeMath As OMath

    Set nativeMath = VTNativeMathNearStart( _
        documentObject, expectedStart, maximumDistance)
    If nativeMath Is Nothing Then
        Err.Raise vbObjectError + 7474, "VisualTeX", _
            "Word could not resolve the native equation after the surrounding document changed."
    End If
    Set VTResolveNativeEquationRange = nativeMath.Range.Duplicate
End Function

Private Function VTOMathCollectionContains( _
    ByVal existingMaths As Collection, _
    ByVal candidateMath As OMath) As Boolean

    Dim existingMath As OMath

    If existingMaths Is Nothing Or candidateMath Is Nothing Then Exit Function
    For Each existingMath In existingMaths
        If existingMath Is candidateMath Then
            VTOMathCollectionContains = True
            Exit Function
        End If
    Next existingMath
End Function

Private Sub VTValidateOmmlFragment(ByVal ommlXml As String)
    Dim normalized As String

    normalized = LTrim$(ommlXml)
    If Len(normalized) = 0 Or Len(normalized) > _
       VT_WORD_OMML_CHUNK_SIZE * VT_WORD_OMML_MAX_CHUNKS Then
        Err.Raise vbObjectError + 7433, "VisualTeX", "The Word OMML payload is empty or too large."
    End If
    If Left$(normalized, 8) <> "<m:oMath" Then
        Err.Raise vbObjectError + 7433, "VisualTeX", "The Word OMML payload must contain one m:oMath root."
    End If
    If InStr(1, normalized, _
        "http:" & "//schemas.openxmlformats.org/officeDocument/2006/math", _
        vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7433, "VisualTeX", "The Word OMML namespace is missing."
    End If
    If InStr(1, normalized, "<!DOCTYPE", vbTextCompare) > 0 Or _
       InStr(1, normalized, "<!ENTITY", vbTextCompare) > 0 Or _
       InStr(1, normalized, "<w:altChunk", vbTextCompare) > 0 Or _
       InStr(1, normalized, "<pkg:package", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 7433, "VisualTeX", "The Word OMML payload contains unsafe XML content."
    End If
End Sub

Private Function VTInsertNativeEquationNumber( _
    ByVal equationRange As Range, _
    ByVal renderedHeightPoints As Double, _
    ByVal formulaId As String, _
    ByVal captionText As String) As Range

    Dim nativeEquation As OMath
    Dim documentObject As Document
    Dim exactEquationRange As Range
    Dim paragraphRange As Range
    Dim insertionRange As Range
    Dim separatorRange As Range
    Dim sequenceField As Field
    Dim equationLabelName As String
    Dim fieldAnchor As Long
    Dim nativeEquationStart As Long
    Dim separatorStart As Long

    If equationRange Is Nothing Or equationRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7470, "VisualTeX", "The native equation number target is missing."
    End If
    Set nativeEquation = equationRange.OMaths(1)
    nativeEquation.Type = wdOMathInline
    Set exactEquationRange = nativeEquation.Range.Duplicate
    nativeEquationStart = exactEquationRange.Start
    Set documentObject = exactEquationRange.Document
    Set paragraphRange = exactEquationRange.Paragraphs(1).Range.Duplicate
    VTConfigureNumberedEquationParagraph paragraphRange
    equationLabelName = VTNativeEquationLabelName()

    Set insertionRange = documentObject.Range( _
        Start:=exactEquationRange.End, End:=exactEquationRange.End)
    insertionRange.Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
    Selection.TypeText Text:=vbTab & "("
    separatorStart = Selection.Start - 2
    Set separatorRange = documentObject.Range( _
        Start:=separatorStart, End:=Selection.Start)
    If separatorRange.Text <> vbTab & "(" Or _
       separatorRange.OMaths.Count <> 0 Or _
       separatorRange.Paragraphs(1).Range.Start <> paragraphRange.Start Then
        separatorRange.Delete
        Err.Raise vbObjectError + 7544, "VisualTeX", _
            "Word could not establish the native Equation number boundary."
    End If
    Set insertionRange = Selection.Range.Duplicate
    insertionRange.Collapse wdCollapseEnd
    Set sequenceField = VTInsertRegisteredEquationCaption( _
        insertionRange, equationLabelName)
    sequenceField.Update
    fieldAnchor = VTEquationFieldStart(sequenceField)

    Set exactEquationRange = VTResolveNativeEquationRange( _
        documentObject, nativeEquationStart, 16)
    VTNormalizeEquationNumberLayoutWithField _
        exactEquationRange, renderedHeightPoints, formulaId, captionText, _
        fieldAnchor
    Set VTInsertNativeEquationNumber = _
        exactEquationRange.Paragraphs(1).Range.Duplicate
End Function

Private Sub VTWordConvertInlineShapeToNativeEquation(ByVal target As InlineShape)
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim ommlBase64 As String
    Dim latexBase64 As String
    Dim captionText As String
    Dim encodedMetadata As String
    Dim nativeDisplayMode As String
    Dim targetDocument As Document
    Dim insertionAnchor As Range
    Dim equationRange As Range
    Dim sourceImage As InlineShape
    Dim sourceBackupDocument As Document
    Dim sourceBackupRange As Range
    Dim sourceRestoreRange As Range
    Dim sourceHeightPoints As Double
    Dim sourceStart As Long
    Dim nativeEquationStart As Long
    Dim sourceDeleted As Boolean
    Dim conversionErrorNumber As Long
    Dim conversionErrorDescription As String
    Dim nativeDocumentPath As String
    Dim numberLayoutRange As Range
    Dim finalFormulaRange As Range
    Dim finalLayoutTable As Table
    Dim nativeBookmarkAnchor As Long
    Dim numberCreated As Boolean
    Dim internalMutationStarted As Boolean

    If target Is Nothing Or Not VTIsVisualTeXInlineShape(target) Then
        Err.Raise vbObjectError + 7430, "VisualTeX", "The selected object is not a VisualTeX formula image."
    End If
    If Not VTTryParseFormulaReference(target.Title, formulaId, displayMode, numbered) Then
        Err.Raise vbObjectError + 7431, "VisualTeX", "The selected VisualTeX formula reference is invalid."
    End If
    Set targetDocument = target.Range.Document
    If Not VTTryReadWordOmmlPayload(targetDocument, formulaId, ommlBase64) Then
        Err.Raise vbObjectError + 7432, "VisualTeX", _
            "This formula has no structural OMML payload. Edit and save it once in the current VisualTeX, then convert it again."
    End If
    encodedMetadata = target.AlternativeText
    If VTTryReadWordLatexPayload(targetDocument, formulaId, latexBase64) Then
        captionText = VTEquationCrossReferenceText(latexBase64)
    Else
        captionText = "VisualTeX formula"
    End If
    sourceHeightPoints = target.Height
    nativeDocumentPath = VTNativeWordDocumentPath(formulaId)
    If Not VTPathFileExists(nativeDocumentPath) Then
        ' Formulas created by an older build have no durable staging DOCX.
        ' Open one edit Session; committing it materializes the cache and
        ' performs the requested conversion instead of showing a dead-end error.
        VTWordEditInlineShape target, True
        Exit Sub
    End If

    nativeDisplayMode = displayMode
    If numbered Or displayMode = "block" Then nativeDisplayMode = "inline"
    sourceStart = target.Range.Start
    Set insertionAnchor = target.Range.Duplicate
    insertionAnchor.Collapse wdCollapseEnd
    On Error GoTo ConversionFailed
    VTBeginWordInternalMutation
    internalMutationStarted = True

    ' Keep an exact FormattedText backup because display promotion must happen
    ' after the source image is removed. Word shifts adjacent OMath Ranges when
    ' an InlineShape is deleted, so the source cannot safely be deleted without
    ' a rollback copy and a fresh equation lookup.
    Set sourceBackupDocument = Documents.Add(Visible:=False)
    Set sourceBackupRange = sourceBackupDocument.Content
    sourceBackupRange.Collapse wdCollapseStart
    sourceBackupRange.FormattedText = target.Range.FormattedText
    If sourceBackupDocument.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7472, "VisualTeX", _
            "Word could not back up the formula image before native conversion."
    End If
    Set sourceBackupRange = _
        sourceBackupDocument.InlineShapes(1).Range.Duplicate
    targetDocument.Activate

    VTSetWordMetadataPayload targetDocument, formulaId, encodedMetadata
    VTSetWordFormulaFormat targetDocument, formulaId, displayMode, numbered
    Set equationRange = VTInsertNativeEquationAtRange( _
        insertionAnchor, _
        ommlBase64, _
        nativeDocumentPath, _
        nativeDisplayMode, _
        displayMode = "block", _
        False)
    nativeEquationStart = equationRange.Start

    targetDocument.Activate
    Set sourceImage = VTFindUniqueInlineShape(encodedMetadata)
    sourceImage.Delete
    sourceDeleted = True

    Set equationRange = VTResolveNativeEquationRange( _
        targetDocument, nativeEquationStart, 16)
    If displayMode = "block" Then
        If numbered Then
            numberCreated = False
            Set numberLayoutRange = VTEnsureNativeEquationNumber( _
                equationRange, sourceHeightPoints, formulaId, captionText, _
                numberCreated)
            Set equationRange = VTResolveNativeEquationRange( _
                targetDocument, nativeEquationStart, 16)
        Else
            Set equationRange = VTPromoteNativeEquationToDisplay(equationRange)
        End If
    ElseIf displayMode = "inline" Then
        Set equationRange = VTFinalizeInlineNativeEquation(equationRange)
    End If
    nativeBookmarkAnchor = equationRange.Start
    VTReconcileEquationNumbers targetDocument

    ' Native SEQ/REF reconciliation can invalidate a Bookmark created from the
    ' pre-refresh OMath Range. Resolve the final formula only after numbering is
    ' stable, then persist VT_F_ as the last structural step of conversion.
    If numbered And displayMode = "block" Then
        If Not targetDocument.Bookmarks.Exists( _
           VTEquationNumberBookmarkName(formulaId)) Then
            Err.Raise vbObjectError + 7460, "VisualTeX", _
                "Word lost the numbered formula identity after native conversion."
        End If
        Set numberLayoutRange = targetDocument.Bookmarks( _
            VTEquationNumberBookmarkName(formulaId)).Range.Duplicate
        If numberLayoutRange.Information(wdWithInTable) Then
            Set numberLayoutRange = numberLayoutRange.Tables(1). _
                Cell(1, 2).Range.Duplicate
            If numberLayoutRange.OMaths.Count <> 1 Then
                Err.Raise vbObjectError + 7460, "VisualTeX", _
                    "Word lost the converted OMath while refreshing its number."
            End If
            Set finalFormulaRange = _
                numberLayoutRange.OMaths(1).Range.Duplicate
        Else
            Set finalFormulaRange = VTResolveNativeEquationRange( _
                targetDocument, nativeBookmarkAnchor, 128)
            If finalFormulaRange.Paragraphs(1).Range.Start <> _
               numberLayoutRange.Paragraphs(1).Range.Start Then
                Err.Raise vbObjectError + 7460, "VisualTeX", _
                    "The converted OMath and its number left the shared paragraph."
            End If
        End If
    Else
        Set finalFormulaRange = VTResolveNativeEquationRange( _
            targetDocument, nativeBookmarkAnchor, 64)
    End If
    Set equationRange = finalFormulaRange.Duplicate

    sourceBackupDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set sourceBackupDocument = Nothing
    On Error Resume Next
    If displayMode = "inline" Then
        VTPlaceCaretAfterInlineNativeEquation equationRange
    Else
        equationRange.Select
    End If
    On Error GoTo 0
    DoEvents

    ' Selection changes and the deferred orphan watcher are allowed to settle
    ' before the final identity transaction. No document or Selection mutation
    ' is permitted after this block: first repair the final visible number
    ' scaffold, then persist VT_F_ over the final center-cell OMath Range.
    Set finalFormulaRange = VTResolveNativeEquationRange( _
        targetDocument, nativeBookmarkAnchor, 128)
    If numbered And displayMode = "block" Then
        If finalFormulaRange.Information(wdWithInTable) Then
            If finalFormulaRange.Tables.Count <> 1 Then
                Err.Raise vbObjectError + 7460, "VisualTeX", _
                    "The converted OMath has an ambiguous numbered table."
            End If
            Set finalLayoutTable = finalFormulaRange.Tables(1)
            VTEnsureEquationNumberFields finalLayoutTable, formulaId
            If Not targetDocument.Bookmarks.Exists( _
               VTEquationNumberBookmarkName(formulaId)) Then
                Err.Raise vbObjectError + 7460, "VisualTeX", _
                    "Word did not restore the final visible Equation identity."
            End If
            Set finalLayoutTable = targetDocument.Bookmarks( _
                VTEquationNumberBookmarkName(formulaId)).Range.Tables(1)
            If finalLayoutTable.Cell(1, 2).Range.OMaths.Count <> 1 Then
                Err.Raise vbObjectError + 7460, "VisualTeX", _
                    "The final numbered table does not contain one converted OMath."
            End If
            Set finalFormulaRange = finalLayoutTable.Cell(1, 2).Range. _
                OMaths(1).Range.Duplicate
        Else
            VTFinalizeParagraphEquationNumber _
                targetDocument, finalFormulaRange, formulaId
            Set finalFormulaRange = VTResolveNativeEquationRange( _
                targetDocument, nativeBookmarkAnchor, 128)
            If finalFormulaRange.Information(wdWithInTable) Then
                Err.Raise vbObjectError + 7460, "VisualTeX", _
                    "The final single-paragraph OMath unexpectedly entered a table."
            End If
        End If
    End If
    VTSetNativeFormulaBookmark _
        targetDocument, finalFormulaRange, formulaId
    If Not targetDocument.Bookmarks.Exists( _
       VTNativeFormulaBookmarkName(formulaId)) Or _
       (numbered And displayMode = "block" And _
        Not targetDocument.Bookmarks.Exists( _
            VTEquationNumberBookmarkName(formulaId))) Then
        Err.Raise vbObjectError + 7460, "VisualTeX", _
            "Word did not preserve the final VisualTeX formula identity."
    End If
    Set equationRange = finalFormulaRange.Duplicate
    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    VTEnsureOrphanWatchScheduled
    Exit Sub

ConversionFailed:
    conversionErrorNumber = Err.Number
    conversionErrorDescription = Err.Description
    On Error Resume Next
    If Not equationRange Is Nothing Then equationRange.Delete
    If sourceDeleted And Not sourceBackupRange Is Nothing Then
        Set sourceRestoreRange = targetDocument.Range( _
            Start:=sourceStart, End:=sourceStart)
        sourceRestoreRange.FormattedText = sourceBackupRange.FormattedText
    End If
    If Not sourceBackupDocument Is Nothing Then
        sourceBackupDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    VTEnsureOrphanWatchScheduled
    On Error GoTo 0
    Err.Raise conversionErrorNumber, "VisualTeX Word image-to-native conversion", conversionErrorDescription
End Sub

Private Function VTNativeWordDocumentPath(ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7434, "VisualTeX", "The native Word formula id is invalid."
    End If
    VTNativeWordDocumentPath = _
        VTApplicationSupportRoot() & "/NativeDocuments/" & formulaId & ".docx"
End Function

Private Function VTNativeFormulaBookmarkName(ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7457, "VisualTeX", "VisualTeX cannot bookmark a native equation with an invalid formula id."
    End If
    VTNativeFormulaBookmarkName = _
        VT_WORD_NATIVE_BOOKMARK_PREFIX & Replace$(formulaId, "-", "")
    If Len(VTNativeFormulaBookmarkName) > 40 Then
        Err.Raise vbObjectError + 7458, "VisualTeX", "VisualTeX generated a native equation bookmark longer than Word permits."
    End If
End Function

Private Function VTTryFormulaIdFromNativeBookmark( _
    ByVal bookmarkName As String, _
    ByRef formulaId As String) As Boolean

    Dim compactId As String

    formulaId = ""
    If Left$(bookmarkName, Len(VT_WORD_NATIVE_BOOKMARK_PREFIX)) <> _
       VT_WORD_NATIVE_BOOKMARK_PREFIX Then Exit Function
    compactId = Mid$(bookmarkName, Len(VT_WORD_NATIVE_BOOKMARK_PREFIX) + 1)
    If Len(compactId) <> 32 Then Exit Function
    formulaId = _
        Left$(compactId, 8) & "-" & _
        Mid$(compactId, 9, 4) & "-" & _
        Mid$(compactId, 13, 4) & "-" & _
        Mid$(compactId, 17, 4) & "-" & _
        Right$(compactId, 12)
    If Not VTIsCanonicalUuid(formulaId) Then
        formulaId = ""
        Exit Function
    End If
    VTTryFormulaIdFromNativeBookmark = True
End Function

Private Sub VTSetNativeFormulaBookmark( _
    ByVal documentObject As Document, _
    ByVal equationRange As Range, _
    ByVal formulaId As String)

    Dim bookmarkName As String
    Dim exactRange As Range
    Dim persistedMath As OMath

    If equationRange Is Nothing Or equationRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7459, "VisualTeX", "VisualTeX cannot bookmark a missing native equation."
    End If
    bookmarkName = VTNativeFormulaBookmarkName(formulaId)
    Set exactRange = equationRange.OMaths(1).Range.Duplicate
    On Error Resume Next
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        documentObject.Bookmarks(bookmarkName).Delete
    End If
    On Error GoTo BookmarkFailed
    documentObject.Bookmarks.Add Name:=bookmarkName, Range:=exactRange
    If Not documentObject.Bookmarks.Exists(bookmarkName) Then GoTo BookmarkFailed
    Set persistedMath = VTNativeMathForBookmark(documentObject.Bookmarks(bookmarkName))
    If persistedMath Is Nothing Then GoTo BookmarkFailed
    Exit Sub

BookmarkFailed:
    Err.Raise vbObjectError + 7460, "VisualTeX", "Word did not persist the VisualTeX native equation bookmark."
End Sub

Private Function VTNativeMathForBookmark(ByVal nativeBookmark As Bookmark) As OMath
    Dim candidate As OMath
    Dim match As OMath
    Dim matchCount As Long
    Dim bookmarkStart As Long
    Dim bookmarkEnd As Long
    Dim candidateDistance As Double
    Dim bestDistance As Double

    If nativeBookmark Is Nothing Then Exit Function
    On Error GoTo NoMatch
    If nativeBookmark.Range.OMaths.Count = 1 Then
        Set VTNativeMathForBookmark = nativeBookmark.Range.OMaths(1)
        Exit Function
    End If

    bookmarkStart = nativeBookmark.Range.Start
    bookmarkEnd = nativeBookmark.Range.End
    bestDistance = 1E+30
    For Each candidate In nativeBookmark.Range.Document.OMaths
        If candidate.Range.Start <= bookmarkEnd + 1 And _
           candidate.Range.End >= bookmarkStart - 1 Then
            candidateDistance = _
                Abs(CDbl(candidate.Range.Start) - CDbl(bookmarkStart)) + _
                Abs(CDbl(candidate.Range.End) - CDbl(bookmarkEnd))
            If candidateDistance < bestDistance Then
                bestDistance = candidateDistance
                matchCount = 1
                Set match = candidate
            ElseIf candidateDistance = bestDistance Then
                matchCount = matchCount + 1
            End If
        End If
    Next candidate
    If matchCount = 1 Then Set VTNativeMathForBookmark = match
    Exit Function

NoMatch:
    Set VTNativeMathForBookmark = Nothing
End Function

Private Function VTFindSelectedNativeFormulaBookmark(ByVal selected As Selection) As Bookmark
    If selected Is Nothing Then
        Err.Raise vbObjectError + 7461, "VisualTeX", "Select one VisualTeX native Word equation."
    End If
    Set VTFindSelectedNativeFormulaBookmark = VTFindNativeFormulaBookmark(selected.Range, True)
End Function

Private Function VTFindNativeFormulaBookmark( _
    ByVal selectedRange As Range, _
    Optional ByVal requireMatch As Boolean = True) As Bookmark

    Dim candidate As Bookmark
    Dim match As Bookmark
    Dim formulaId As String
    Dim matchCount As Long
    Dim candidateMath As OMath

    If selectedRange Is Nothing Then GoTo NoMatch
    For Each candidate In selectedRange.Document.Bookmarks
        If VTTryFormulaIdFromNativeBookmark(candidate.Name, formulaId) Then
            Set candidateMath = VTNativeMathForBookmark(candidate)
            If Not candidateMath Is Nothing Then
                If selectedRange.Start <= candidate.Range.End And _
                   selectedRange.End >= candidate.Range.Start Then
                    matchCount = matchCount + 1
                    Set match = candidate
                End If
            End If
        End If
    Next candidate
    If matchCount = 1 Then
        Set VTFindNativeFormulaBookmark = match
        Exit Function
    End If
    If matchCount > 1 Then
        Err.Raise vbObjectError + 7462, "VisualTeX", "The selection intersects multiple VisualTeX native equations."
    End If

NoMatch:
    If requireMatch Then
        Err.Raise vbObjectError + 7461, "VisualTeX", "Select one VisualTeX formula image or native equation."
    End If
End Function

Private Function VTWordLatexVariableStem(ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7435, "VisualTeX", "VisualTeX cannot address Word LaTeX metadata for an invalid formula id."
    End If
    VTWordLatexVariableStem = VT_WORD_LATEX_VARIABLE_PREFIX & Replace$(formulaId, "-", "_")
End Function

Private Function VTWordLatexCountVariableName(ByVal formulaId As String) As String
    VTWordLatexCountVariableName = VTWordLatexVariableStem(formulaId) & "_Count"
End Function

Private Function VTWordLatexChunkVariableName(ByVal formulaId As String, ByVal index As Long) As String
    If index < 1 Or index > VT_WORD_LATEX_MAX_CHUNKS Then
        Err.Raise vbObjectError + 7436, "VisualTeX", "VisualTeX Word LaTeX metadata chunk index is invalid."
    End If
    VTWordLatexChunkVariableName = _
        VTWordLatexVariableStem(formulaId) & "_" & Right$("000" & CStr(index), 3)
End Function

Private Function VTIsBase64UrlPayload(ByVal value As String) As Boolean
    Dim index As Long
    Dim current As String

    If Len(value) = 0 Or Len(value) > VT_WORD_LATEX_CHUNK_SIZE * VT_WORD_LATEX_MAX_CHUNKS Then Exit Function
    If Len(value) Mod 4 = 1 Then Exit Function
    For index = 1 To Len(value)
        current = Mid$(value, index, 1)
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_", current, vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next index
    VTIsBase64UrlPayload = True
End Function

Private Function VTTryGetDocumentVariable( _
    ByVal documentObject As Document, _
    ByVal variableName As String, _
    ByRef variableValue As String) As Boolean

    On Error Resume Next
    Err.Clear
    variableValue = documentObject.Variables(variableName).Value
    VTTryGetDocumentVariable = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub VTSetDocumentVariable( _
    ByVal documentObject As Document, _
    ByVal variableName As String, _
    ByVal variableValue As String)

    Dim ignored As String
    If VTTryGetDocumentVariable(documentObject, variableName, ignored) Then
        documentObject.Variables(variableName).Value = variableValue
    Else
        documentObject.Variables.Add Name:=variableName, Value:=variableValue
    End If
End Sub

Private Sub VTDeleteDocumentVariable(ByVal documentObject As Document, ByVal variableName As String)
    On Error Resume Next
    documentObject.Variables(variableName).Delete
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub VTDeleteWordLatexPayload(ByVal documentObject As Document, ByVal formulaId As String)
    Dim index As Long

    For index = 1 To VT_WORD_LATEX_MAX_CHUNKS
        VTDeleteDocumentVariable documentObject, VTWordLatexChunkVariableName(formulaId, index)
    Next index
    VTDeleteDocumentVariable documentObject, VTWordLatexCountVariableName(formulaId)
End Sub

Private Sub VTSetWordLatexPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal latexBase64 As String)

    Dim chunkCount As Long
    Dim index As Long
    Dim chunkValue As String
    Dim storageErrorNumber As Long
    Dim storageErrorDescription As String

    If Not VTIsBase64UrlPayload(latexBase64) Then
        Err.Raise vbObjectError + 7437, "VisualTeX", "VisualTeX Word LaTeX metadata is invalid or too large."
    End If
    chunkCount = (Len(latexBase64) + VT_WORD_LATEX_CHUNK_SIZE - 1) \ VT_WORD_LATEX_CHUNK_SIZE
    If chunkCount < 1 Or chunkCount > VT_WORD_LATEX_MAX_CHUNKS Then
        Err.Raise vbObjectError + 7437, "VisualTeX", "VisualTeX Word LaTeX metadata requires too many chunks."
    End If

    VTDeleteWordLatexPayload documentObject, formulaId
    On Error GoTo StorageFailed
    For index = 1 To chunkCount
        chunkValue = Mid$( _
            latexBase64, _
            (index - 1) * VT_WORD_LATEX_CHUNK_SIZE + 1, _
            VT_WORD_LATEX_CHUNK_SIZE)
        VTSetDocumentVariable _
            documentObject, _
            VTWordLatexChunkVariableName(formulaId, index), _
            chunkValue
    Next index
    ' Publish the count last so readers never accept a partially written payload.
    VTSetDocumentVariable _
        documentObject, _
        VTWordLatexCountVariableName(formulaId), _
        CStr(chunkCount)
    Exit Sub

StorageFailed:
    storageErrorNumber = Err.Number
    storageErrorDescription = Err.Description
    On Error Resume Next
    VTDeleteWordLatexPayload documentObject, formulaId
    On Error GoTo 0
    Err.Raise storageErrorNumber, "VisualTeX Word LaTeX metadata", storageErrorDescription
End Sub

Private Function VTTryReadWordLatexPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByRef latexBase64 As String) As Boolean

    Dim countText As String
    Dim chunkValue As String
    Dim chunkCount As Long
    Dim index As Long

    latexBase64 = ""
    If Not VTTryGetDocumentVariable( _
        documentObject, VTWordLatexCountVariableName(formulaId), countText) Then Exit Function
    If Len(countText) = 0 Or Not IsNumeric(countText) Then GoTo InvalidPayload
    chunkCount = CLng(countText)
    If chunkCount < 1 Or chunkCount > VT_WORD_LATEX_MAX_CHUNKS Then GoTo InvalidPayload

    For index = 1 To chunkCount
        If Not VTTryGetDocumentVariable( _
            documentObject, _
            VTWordLatexChunkVariableName(formulaId, index), _
            chunkValue) Then GoTo InvalidPayload
        latexBase64 = latexBase64 & chunkValue
    Next index
    If Not VTIsBase64UrlPayload(latexBase64) Then GoTo InvalidPayload
    VTTryReadWordLatexPayload = True
    Exit Function

InvalidPayload:
    Err.Raise vbObjectError + 7438, "VisualTeX", "The stored Word native-equation LaTeX metadata is incomplete or corrupt."
End Function

Private Function VTWordOmmlVariableStem(ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7471, "VisualTeX", "VisualTeX cannot address Word OMML metadata for an invalid formula id."
    End If
    VTWordOmmlVariableStem = _
        VT_WORD_OMML_VARIABLE_PREFIX & Replace$(formulaId, "-", "_")
End Function

Private Function VTWordOmmlCountVariableName(ByVal formulaId As String) As String
    VTWordOmmlCountVariableName = VTWordOmmlVariableStem(formulaId) & "_Count"
End Function

Private Function VTWordOmmlChunkVariableName( _
    ByVal formulaId As String, _
    ByVal index As Long) As String

    If index < 1 Or index > VT_WORD_OMML_MAX_CHUNKS Then
        Err.Raise vbObjectError + 7472, "VisualTeX", "VisualTeX Word OMML metadata chunk index is invalid."
    End If
    VTWordOmmlChunkVariableName = _
        VTWordOmmlVariableStem(formulaId) & "_" & Right$("000" & CStr(index), 3)
End Function

Private Sub VTDeleteWordOmmlPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String)

    Dim index As Long
    For index = 1 To VT_WORD_OMML_MAX_CHUNKS
        VTDeleteDocumentVariable _
            documentObject, VTWordOmmlChunkVariableName(formulaId, index)
    Next index
    VTDeleteDocumentVariable documentObject, VTWordOmmlCountVariableName(formulaId)
End Sub

Private Sub VTSetWordOmmlPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal ommlBase64 As String)

    Dim chunkCount As Long
    Dim index As Long
    Dim chunkValue As String
    Dim storageErrorNumber As Long
    Dim storageErrorDescription As String

    If Not VTIsBase64UrlPayload(ommlBase64) Then
        Err.Raise vbObjectError + 7473, "VisualTeX", "VisualTeX Word OMML metadata is invalid or too large."
    End If
    chunkCount = _
        (Len(ommlBase64) + VT_WORD_OMML_CHUNK_SIZE - 1) \ _
        VT_WORD_OMML_CHUNK_SIZE
    If chunkCount < 1 Or chunkCount > VT_WORD_OMML_MAX_CHUNKS Then
        Err.Raise vbObjectError + 7473, "VisualTeX", "VisualTeX Word OMML metadata requires too many chunks."
    End If

    VTDeleteWordOmmlPayload documentObject, formulaId
    On Error GoTo StorageFailed
    For index = 1 To chunkCount
        chunkValue = Mid$( _
            ommlBase64, _
            (index - 1) * VT_WORD_OMML_CHUNK_SIZE + 1, _
            VT_WORD_OMML_CHUNK_SIZE)
        VTSetDocumentVariable _
            documentObject, _
            VTWordOmmlChunkVariableName(formulaId, index), _
            chunkValue
    Next index
    VTSetDocumentVariable _
        documentObject, _
        VTWordOmmlCountVariableName(formulaId), _
        CStr(chunkCount)
    Exit Sub

StorageFailed:
    storageErrorNumber = Err.Number
    storageErrorDescription = Err.Description
    On Error Resume Next
    VTDeleteWordOmmlPayload documentObject, formulaId
    On Error GoTo 0
    Err.Raise storageErrorNumber, "VisualTeX Word OMML metadata", storageErrorDescription
End Sub

Private Function VTTryReadWordOmmlPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByRef ommlBase64 As String) As Boolean

    Dim countText As String
    Dim chunkValue As String
    Dim chunkCount As Long
    Dim index As Long
    Dim ommlXml As String

    ommlBase64 = ""
    If Not VTTryGetDocumentVariable( _
        documentObject, VTWordOmmlCountVariableName(formulaId), countText) Then Exit Function
    If Len(countText) = 0 Or Not IsNumeric(countText) Then GoTo InvalidPayload
    chunkCount = CLng(countText)
    If chunkCount < 1 Or chunkCount > VT_WORD_OMML_MAX_CHUNKS Then GoTo InvalidPayload

    For index = 1 To chunkCount
        If Not VTTryGetDocumentVariable( _
            documentObject, _
            VTWordOmmlChunkVariableName(formulaId, index), _
            chunkValue) Then GoTo InvalidPayload
        ommlBase64 = ommlBase64 & chunkValue
    Next index
    If Not VTIsBase64UrlPayload(ommlBase64) Then GoTo InvalidPayload
    ommlXml = VTBase64UrlDecodeUtf8(ommlBase64)
    VTValidateOmmlFragment ommlXml
    VTTryReadWordOmmlPayload = True
    Exit Function

InvalidPayload:
    Err.Raise vbObjectError + 7474, "VisualTeX", "The stored Word OMML metadata is incomplete or corrupt."
End Function

Private Function VTWordMetadataVariableStem(ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7463, "VisualTeX", "VisualTeX cannot address Word formula metadata for an invalid formula id."
    End If
    VTWordMetadataVariableStem = _
        VT_WORD_METADATA_VARIABLE_PREFIX & Replace$(formulaId, "-", "_")
End Function

Private Function VTWordMetadataCountVariableName(ByVal formulaId As String) As String
    VTWordMetadataCountVariableName = VTWordMetadataVariableStem(formulaId) & "_Count"
End Function

Private Function VTWordMetadataChunkVariableName( _
    ByVal formulaId As String, _
    ByVal index As Long) As String

    If index < 1 Or index > VT_WORD_PAYLOAD_MAX_CHUNKS Then
        Err.Raise vbObjectError + 7464, "VisualTeX", "VisualTeX Word metadata chunk index is invalid."
    End If
    VTWordMetadataChunkVariableName = _
        VTWordMetadataVariableStem(formulaId) & "_" & Right$("000" & CStr(index), 3)
End Function

Private Sub VTDeleteWordMetadataPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String)

    Dim index As Long
    For index = 1 To VT_WORD_PAYLOAD_MAX_CHUNKS
        VTDeleteDocumentVariable _
            documentObject, VTWordMetadataChunkVariableName(formulaId, index)
    Next index
    VTDeleteDocumentVariable documentObject, VTWordMetadataCountVariableName(formulaId)
End Sub

Private Sub VTSetWordMetadataPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal encodedMetadata As String)

    Dim chunkCount As Long
    Dim index As Long
    Dim chunkValue As String
    Dim storageErrorNumber As Long
    Dim storageErrorDescription As String

    If Not VTIsEncodedMetadata(encodedMetadata) Then
        Err.Raise vbObjectError + 7465, "VisualTeX", "VisualTeX Word formula metadata is invalid or too large."
    End If
    chunkCount = _
        (Len(encodedMetadata) + VT_WORD_PAYLOAD_CHUNK_SIZE - 1) \ _
        VT_WORD_PAYLOAD_CHUNK_SIZE
    If chunkCount < 1 Or chunkCount > VT_WORD_PAYLOAD_MAX_CHUNKS Then
        Err.Raise vbObjectError + 7465, "VisualTeX", "VisualTeX Word formula metadata requires too many chunks."
    End If

    VTDeleteWordMetadataPayload documentObject, formulaId
    On Error GoTo StorageFailed
    For index = 1 To chunkCount
        chunkValue = Mid$( _
            encodedMetadata, _
            (index - 1) * VT_WORD_PAYLOAD_CHUNK_SIZE + 1, _
            VT_WORD_PAYLOAD_CHUNK_SIZE)
        VTSetDocumentVariable _
            documentObject, _
            VTWordMetadataChunkVariableName(formulaId, index), _
            chunkValue
    Next index
    VTSetDocumentVariable _
        documentObject, _
        VTWordMetadataCountVariableName(formulaId), _
        CStr(chunkCount)
    Exit Sub

StorageFailed:
    storageErrorNumber = Err.Number
    storageErrorDescription = Err.Description
    On Error Resume Next
    VTDeleteWordMetadataPayload documentObject, formulaId
    On Error GoTo 0
    Err.Raise storageErrorNumber, "VisualTeX Word formula metadata", storageErrorDescription
End Sub

Private Function VTTryReadWordMetadataPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByRef encodedMetadata As String) As Boolean

    Dim countText As String
    Dim chunkValue As String
    Dim chunkCount As Long
    Dim index As Long

    encodedMetadata = ""
    If Not VTTryGetDocumentVariable( _
        documentObject, VTWordMetadataCountVariableName(formulaId), countText) Then Exit Function
    If Len(countText) = 0 Or Not IsNumeric(countText) Then GoTo InvalidPayload
    chunkCount = CLng(countText)
    If chunkCount < 1 Or chunkCount > VT_WORD_PAYLOAD_MAX_CHUNKS Then GoTo InvalidPayload

    For index = 1 To chunkCount
        If Not VTTryGetDocumentVariable( _
            documentObject, _
            VTWordMetadataChunkVariableName(formulaId, index), _
            chunkValue) Then GoTo InvalidPayload
        encodedMetadata = encodedMetadata & chunkValue
    Next index
    If Not VTIsEncodedMetadata(encodedMetadata) Then GoTo InvalidPayload
    VTTryReadWordMetadataPayload = True
    Exit Function

InvalidPayload:
    Err.Raise vbObjectError + 7466, "VisualTeX", "The stored Word formula metadata is incomplete or corrupt."
End Function

Private Function VTWordFormatVariableName(ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7467, "VisualTeX", "VisualTeX cannot address Word formula format for an invalid formula id."
    End If
    VTWordFormatVariableName = _
        VT_WORD_FORMAT_VARIABLE_PREFIX & Replace$(formulaId, "-", "_")
End Function

Private Sub VTSetWordFormulaFormat( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal displayMode As String, _
    ByVal numbered As Boolean)

    If displayMode <> "inline" And displayMode <> "block" Then
        Err.Raise vbObjectError + 7468, "VisualTeX", "VisualTeX Word formula format has an invalid display mode."
    End If
    If numbered And displayMode <> "block" Then
        Err.Raise vbObjectError + 7468, "VisualTeX", "Only display formulas can retain Word equation numbers."
    End If
    VTSetDocumentVariable _
        documentObject, _
        VTWordFormatVariableName(formulaId), _
        displayMode & "|" & IIf(numbered, "1", "0")
End Sub

Private Function VTTryReadWordFormulaFormat( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByRef displayMode As String, _
    ByRef numbered As Boolean) As Boolean

    Dim storedValue As String
    Dim fields() As String

    displayMode = ""
    numbered = False
    If Not VTTryGetDocumentVariable( _
        documentObject, VTWordFormatVariableName(formulaId), storedValue) Then Exit Function
    fields = Split(storedValue, "|")
    If UBound(fields) <> 1 Then GoTo InvalidFormat
    displayMode = fields(0)
    If displayMode <> "inline" And displayMode <> "block" Then GoTo InvalidFormat
    If fields(1) = "1" Then
        numbered = True
    ElseIf fields(1) <> "0" Then
        GoTo InvalidFormat
    End If
    If numbered And displayMode <> "block" Then GoTo InvalidFormat
    VTTryReadWordFormulaFormat = True
    Exit Function

InvalidFormat:
    Err.Raise vbObjectError + 7469, "VisualTeX", "The stored Word formula format is invalid or corrupt."
End Function

Private Function VTLaTeXToWordLinear(ByVal latex As String) As String
    Dim normalized As String
    Dim position As Long
    Dim converted As String
    Dim hasMultipleRows As Boolean

    If Len(latex) = 0 Or Len(latex) > 1048576 Then
        Err.Raise vbObjectError + 7439, "VisualTeX", "The stored LaTeX is empty or too large to convert."
    End If

    normalized = VTNormalizeLatexForWord(latex)
    hasMultipleRows = _
        InStr(1, normalized, vbLf, vbBinaryCompare) > 0 Or _
        InStr(1, normalized, "\\", vbBinaryCompare) > 0
    position = 1
    converted = VTConvertLatexSegment(normalized, position, "", 0)
    If position <= Len(normalized) Then
        Err.Raise vbObjectError + 7440, "VisualTeX", "VisualTeX did not consume the complete LaTeX expression."
    End If
    converted = Trim$(converted)
    If hasMultipleRows Then
        If Left$(LTrim$(converted), 8) <> "\matrix(" And _
           Left$(LTrim$(converted), 7) <> "\cases(" Then
            converted = "\matrix(" & converted & ")"
        End If
    End If
    VTLaTeXToWordLinear = converted
End Function

Private Function VTNormalizeLatexForWord(ByVal latex As String) As String
    Dim result As String

    result = Replace$(latex, vbCrLf, vbLf)
    result = Replace$(result, vbCr, vbLf)

    result = Replace$(result, "\begin{pmatrix}", "(\matrix(")
    result = Replace$(result, "\end{pmatrix}", "))")
    result = Replace$(result, "\begin{bmatrix}", "[\matrix(")
    result = Replace$(result, "\end{bmatrix}", ")]")
    result = Replace$(result, "\begin{Bmatrix}", "{\matrix(")
    result = Replace$(result, "\end{Bmatrix}", ")}")
    result = Replace$(result, "\begin{vmatrix}", "|\matrix(")
    result = Replace$(result, "\end{vmatrix}", ")|")
    result = Replace$(result, "\begin{Vmatrix}", "\Vert\matrix(")
    result = Replace$(result, "\end{Vmatrix}", ")\Vert")
    result = Replace$(result, "\begin{matrix}", "\matrix(")
    result = Replace$(result, "\end{matrix}", ")")
    result = Replace$(result, "\begin{cases}", "\cases(")
    result = Replace$(result, "\end{cases}", ")")
    result = Replace$(result, "\begin{aligned}", "\matrix(")
    result = Replace$(result, "\end{aligned}", ")")
    result = Replace$(result, "\begin{alignedat}", "\matrix(")
    result = Replace$(result, "\end{alignedat}", ")")
    result = Replace$(result, "\begin{gathered}", "\matrix(")
    result = Replace$(result, "\end{gathered}", ")")
    result = Replace$(result, "\begin{split}", "\matrix(")
    result = Replace$(result, "\end{split}", ")")
    result = Replace$(result, "\begin{align}", "\matrix(")
    result = Replace$(result, "\end{align}", ")")
    result = Replace$(result, "\begin{align*}", "\matrix(")
    result = Replace$(result, "\end{align*}", ")")
    result = Replace$(result, "\begin{gather}", "\matrix(")
    result = Replace$(result, "\end{gather}", ")")
    result = Replace$(result, "\begin{gather*}", "\matrix(")
    result = Replace$(result, "\end{gather*}", ")")
    result = Replace$(result, "\begin{equation}", "")
    result = Replace$(result, "\end{equation}", "")
    result = Replace$(result, "\begin{equation*}", "")
    result = Replace$(result, "\end{equation*}", "")

    result = Trim$(result)
    If Len(result) >= 2 And Left$(result, 1) = "$" And Right$(result, 1) = "$" Then
        result = Mid$(result, 2, Len(result) - 2)
    End If
    If Left$(result, 2) = "\(" And Right$(result, 2) = "\)" Then
        result = Mid$(result, 3, Len(result) - 4)
    ElseIf Left$(result, 2) = "\[" And Right$(result, 2) = "\]" Then
        result = Mid$(result, 3, Len(result) - 4)
    End If
    VTNormalizeLatexForWord = result
End Function

Private Function VTConvertLatexSegment( _
    ByVal source As String, _
    ByRef position As Long, _
    ByVal terminator As String, _
    ByVal depth As Long) As String

    Dim result As String
    Dim current As String
    Dim atom As String

    If depth > 64 Then
        Err.Raise vbObjectError + 7441, "VisualTeX", "The LaTeX expression is nested too deeply for Word conversion."
    End If

    Do While position <= Len(source)
        current = Mid$(source, position, 1)
        If Len(terminator) > 0 And current = terminator Then
            position = position + 1
            VTConvertLatexSegment = result
            Exit Function
        End If

        Select Case current
            Case "\"
                result = result & VTConvertLatexCommand(source, position, depth + 1)
            Case "{"
                position = position + 1
                atom = VTConvertLatexSegment(source, position, "}", depth + 1)
                result = result & "(" & atom & ")"
            Case "}"
                Err.Raise vbObjectError + 7442, "VisualTeX", "The LaTeX expression contains an unmatched closing brace."
            Case "^", "_"
                position = position + 1
                atom = VTReadLatexAtom(source, position, depth + 1)
                result = result & current & "(" & atom & ")"
            Case vbLf
                result = result & "@"
                position = position + 1
            Case Else
                result = result & current
                position = position + 1
        End Select
    Loop

    If Len(terminator) > 0 Then
        Err.Raise vbObjectError + 7443, "VisualTeX", "The LaTeX expression contains an unclosed group."
    End If
    VTConvertLatexSegment = result
End Function

Private Function VTConvertLatexCommand( _
    ByVal source As String, _
    ByRef position As Long, _
    ByVal depth As Long) As String

    Dim commandName As String
    Dim lowerName As String
    Dim firstArgument As String
    Dim secondArgument As String
    Dim optionalArgument As String

    commandName = VTReadLatexCommand(source, position)
    lowerName = LCase$(commandName)

    Select Case lowerName
        Case "\"
            VTConvertLatexCommand = "@"
        Case "frac", "dfrac", "tfrac"
            firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            secondArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            ' Word OMath.BuildUp consumes UnicodeMath, not LaTeX. A true
            ' stacked fraction is written as a grouped numerator divided by a
            ' grouped denominator; the LaTeX-like \frac(a,b) form instead
            ' attaches a malformed denominator to the preceding atom.
            VTConvertLatexCommand = "(" & firstArgument & ")/(" & secondArgument & ")"
        Case "sqrt"
            optionalArgument = VTReadOptionalLatexBracket(source, position, depth + 1)
            firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            If Len(optionalArgument) > 0 Then
                VTConvertLatexCommand = "\sqrt(" & optionalArgument & "&" & firstArgument & ")"
            Else
                VTConvertLatexCommand = "\sqrt(" & firstArgument & ")"
            End If
        Case "binom", "dbinom", "tbinom"
            firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            secondArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            VTConvertLatexCommand = "\binom(" & firstArgument & "&" & secondArgument & ")"
        Case "overset", "stackrel"
            firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            secondArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            VTConvertLatexCommand = "(" & secondArgument & ")^(" & firstArgument & ")"
        Case "underset"
            firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            secondArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            VTConvertLatexCommand = "(" & secondArgument & ")_(" & firstArgument & ")"
        Case "text", "textrm", "textnormal", "operatorname"
            firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            VTConvertLatexCommand = """" & Replace$(firstArgument, """", """""") & """"
        Case "mathrm", "mathbf", "mathit", "mathsf", "mathtt", "mathcal", "mathbb", "mathfrak", _
             "boldsymbol", "bm", "displaystyle", "textstyle", "scriptstyle", "scriptscriptstyle"
            If lowerName = "displaystyle" Or lowerName = "textstyle" Or _
               lowerName = "scriptstyle" Or lowerName = "scriptscriptstyle" Then
                VTConvertLatexCommand = ""
            Else
                firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
                VTConvertLatexCommand = "(" & firstArgument & ")"
            End If
        Case "overline", "bar"
            firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            VTConvertLatexCommand = "\bar(" & firstArgument & ")"
        Case "underline", "underbar"
            firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            VTConvertLatexCommand = "\underbar(" & firstArgument & ")"
        Case "hat", "widehat", "tilde", "widetilde", "vec", "dot", "ddot", "breve", "check", "acute", "grave"
            firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            VTConvertLatexCommand = "\" & commandName & "(" & firstArgument & ")"
        Case "substack"
            firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            VTConvertLatexCommand = "\matrix(" & firstArgument & ")"
        Case "left", "right"
            VTConvertLatexCommand = VTReadLatexDelimiter(source, position)
        Case "label", "tag"
            firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            VTConvertLatexCommand = ""
        Case "nonumber", "notag", "limits", "nolimits"
            VTConvertLatexCommand = ""
        Case "phantom", "hphantom", "vphantom", "boxed"
            firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            VTConvertLatexCommand = "(" & firstArgument & ")"
        Case "textcolor"
            firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            secondArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            VTConvertLatexCommand = "(" & secondArgument & ")"
        Case "color"
            firstArgument = VTReadRequiredLatexGroup(source, position, depth + 1)
            VTConvertLatexCommand = ""
        Case "begin", "end"
            Err.Raise vbObjectError + 7444, "VisualTeX", _
                "This LaTeX environment is not yet supported by native Word equation conversion."
        Case ",", ";", ":", "!", " ", "quad", "qquad", "enspace", "thinspace"
            VTConvertLatexCommand = " "
        Case "{", "}", "_", "%", "#", "&", "$"
            VTConvertLatexCommand = commandName
        Case Else
            ' Word's UnicodeMath parser natively recognizes Greek letters,
            ' large operators, arrows, relations and named functions written
            ' with the same backslash command names used by LaTeX.
            VTConvertLatexCommand = "\" & commandName
    End Select
End Function

Private Function VTReadLatexCommand(ByVal source As String, ByRef position As Long) As String
    Dim startPosition As Long
    Dim current As String

    If Mid$(source, position, 1) <> "\" Then
        Err.Raise vbObjectError + 7445, "VisualTeX", "VisualTeX expected a LaTeX command."
    End If
    position = position + 1
    If position > Len(source) Then
        Err.Raise vbObjectError + 7445, "VisualTeX", "The LaTeX expression ends with an incomplete command."
    End If

    current = Mid$(source, position, 1)
    If current Like "[A-Za-z]" Then
        startPosition = position
        Do While position <= Len(source) And Mid$(source, position, 1) Like "[A-Za-z]"
            position = position + 1
        Loop
        VTReadLatexCommand = Mid$(source, startPosition, position - startPosition)
        If position <= Len(source) And Mid$(source, position, 1) = "*" Then
            VTReadLatexCommand = VTReadLatexCommand & "*"
            position = position + 1
        End If
    Else
        VTReadLatexCommand = current
        position = position + 1
    End If
End Function

Private Function VTReadRequiredLatexGroup( _
    ByVal source As String, _
    ByRef position As Long, _
    ByVal depth As Long) As String

    VTSkipLatexSpaces source, position
    If position > Len(source) Or Mid$(source, position, 1) <> "{" Then
        Err.Raise vbObjectError + 7446, "VisualTeX", "A required LaTeX command argument is missing."
    End If
    position = position + 1
    VTReadRequiredLatexGroup = VTConvertLatexSegment(source, position, "}", depth + 1)
End Function

Private Function VTReadOptionalLatexBracket( _
    ByVal source As String, _
    ByRef position As Long, _
    ByVal depth As Long) As String

    VTSkipLatexSpaces source, position
    If position <= Len(source) And Mid$(source, position, 1) = "[" Then
        position = position + 1
        VTReadOptionalLatexBracket = VTConvertLatexSegment(source, position, "]", depth + 1)
    End If
End Function

Private Function VTReadLatexAtom( _
    ByVal source As String, _
    ByRef position As Long, _
    ByVal depth As Long) As String

    VTSkipLatexSpaces source, position
    If position > Len(source) Then
        Err.Raise vbObjectError + 7447, "VisualTeX", "A LaTeX superscript or subscript argument is missing."
    End If

    Select Case Mid$(source, position, 1)
        Case "{"
            position = position + 1
            VTReadLatexAtom = VTConvertLatexSegment(source, position, "}", depth + 1)
        Case "\"
            VTReadLatexAtom = VTConvertLatexCommand(source, position, depth + 1)
        Case Else
            VTReadLatexAtom = Mid$(source, position, 1)
            position = position + 1
    End Select
End Function

Private Function VTReadLatexDelimiter(ByVal source As String, ByRef position As Long) As String
    Dim commandName As String
    Dim current As String

    VTSkipLatexSpaces source, position
    If position > Len(source) Then Exit Function
    current = Mid$(source, position, 1)
    If current = "\" Then
        commandName = VTReadLatexCommand(source, position)
        Select Case commandName
            Case ".": VTReadLatexDelimiter = ""
            Case "{", "}": VTReadLatexDelimiter = commandName
            Case Else: VTReadLatexDelimiter = "\" & commandName
        End Select
    Else
        position = position + 1
        If current <> "." Then VTReadLatexDelimiter = current
    End If
End Function

Private Sub VTSkipLatexSpaces(ByVal source As String, ByRef position As Long)
    Do While position <= Len(source)
        If Mid$(source, position, 1) <> " " And Mid$(source, position, 1) <> vbTab Then Exit Do
        position = position + 1
    Loop
End Sub

Private Sub VTRequireWritableWordDocument()
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7411, "VisualTeX", "Open a Word document first."
    End If
    If ActiveDocument.ReadOnly Then
        Err.Raise vbObjectError + 7412, "VisualTeX", "The active Word document is read-only."
    End If
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        Err.Raise vbObjectError + 7413, "VisualTeX", "The active Word document is protected."
    End If
End Sub

Private Function VTWordDocumentIdentity() As String
    On Error Resume Next
    VTWordDocumentIdentity = ActiveDocument.FullName
    If Err.Number <> 0 Or Len(VTWordDocumentIdentity) = 0 Then
        Err.Clear
        VTWordDocumentIdentity = ActiveDocument.Name
    End If
    On Error GoTo 0
    VTWordDocumentIdentity = VTBoundedIdentity(VTWordDocumentIdentity)
End Function

Private Sub VTTraceWordSession( _
    ByVal sessionId As String, _
    ByVal eventName As String, _
    ByVal pendingMarker As String)

    Dim tracePath As String
    Dim traceText As String
    Dim item As InlineShape
    Dim index As Long
    Dim bookmarkName As String

    If Not VT_WORD_TRACE_ENABLED Then Exit Sub

    On Error Resume Next
    tracePath = VTSessionDirectory(sessionId) & "/word-trace.log"
    If VTPathFileExists(tracePath) Then traceText = VTReadText(tracePath, 524288)
    bookmarkName = VTWordBookmarkName(sessionId)
    traceText = traceText & _
        "event=" & eventName & _
        " time=" & CStr(Now) & _
        " documentId=" & VTWordDocumentIdentity() & _
        " documents=" & CStr(Documents.Count) & _
        " inlineShapes=" & CStr(ActiveDocument.InlineShapes.Count) & _
        " bookmark=" & bookmarkName & _
        " bookmarkExists=" & CStr(ActiveDocument.Bookmarks.Exists(bookmarkName)) & _
        " marker=" & pendingMarker & vbLf
    For Each item In ActiveDocument.InlineShapes
        index = index + 1
        traceText = traceText & _
            "shape=" & CStr(index) & _
            " start=" & CStr(item.Range.Start) & _
            " end=" & CStr(item.Range.End) & _
            " width=" & CStr(item.Width) & _
            " height=" & CStr(item.Height) & _
            " title=" & item.Title & _
            " alt=" & item.AlternativeText & vbLf
    Next item
    VTWriteTextAtomic tracePath, traceText
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub VTWriteWordFailureTrace( _
    ByVal sessionId As String, _
    ByVal transactionStage As String, _
    ByVal errorNumber As Long, _
    ByVal errorDescription As String)

    Dim tracePath As String
    Dim traceText As String

    On Error Resume Next
    tracePath = VTSessionDirectory(sessionId) & "/word-failure.log"
    traceText = _
        "time=" & CStr(Now) & vbLf & _
        "stage=" & transactionStage & vbLf & _
        "errorNumber=" & CStr(errorNumber) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            errorDescription, vbCr, " "), vbLf, " ") & vbLf & _
        "documentId=" & VTWordDocumentIdentity() & vbLf
    VTWriteTextAtomic tracePath, traceText
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub VTAddPendingBookmark(ByVal targetRange As Range, ByVal sessionId As String)
    Dim name As String
    name = VTWordBookmarkName(sessionId)
    On Error Resume Next
    If ActiveDocument.Bookmarks.Exists(name) Then ActiveDocument.Bookmarks(name).Delete
    On Error GoTo 0
    ActiveDocument.Bookmarks.Add Name:=name, Range:=targetRange
End Sub

Private Sub VTDeletePendingBookmark( _
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

Private Function VTWordBookmarkName(ByVal sessionId As String) As String
    If Not VTIsCanonicalUuid(sessionId) Then
        Err.Raise vbObjectError + 7420, "VisualTeX", "VisualTeX cannot create a Bookmark for an invalid Session id."
    End If
    VTWordBookmarkName = VT_WORD_BOOKMARK_PREFIX & Replace$(Left$(sessionId, 24), "-", "")
    If Len(VTWordBookmarkName) > 40 Then
        Err.Raise vbObjectError + 7421, "VisualTeX", "VisualTeX generated a Word Bookmark name longer than 40 characters."
    End If
End Function

Private Function VTDispatchOptional(ByVal dispatch As Object, ByVal key As String) As String
    If VTCollectionHasKey(dispatch, key) Then VTDispatchOptional = CStr(dispatch(key))
End Function

Private Function VTDispatchPositiveDouble(ByVal dispatch As Object, ByVal key As String) As Double
    VTRequireDispatchValue dispatch, key
    VTDispatchPositiveDouble = VTParseInvariantDouble(CStr(dispatch(key)))
    If VTDispatchPositiveDouble <= 0# Or VTDispatchPositiveDouble > 100000# Then
        Err.Raise vbObjectError + 7414, "VisualTeX", "VisualTeX dispatch contains invalid " & key & "."
    End If
End Function

Private Function VTDispatchOptionalDouble(ByVal dispatch As Object, ByVal key As String, ByVal fallback As Double) As Double
    If Not VTCollectionHasKey(dispatch, key) Or Len(CStr(dispatch(key))) = 0 Then
        VTDispatchOptionalDouble = fallback
    Else
        VTDispatchOptionalDouble = VTParseInvariantDouble(CStr(dispatch(key)))
    End If
End Function

Private Sub VTRefreshWordHealthQuietly()
    On Error Resume Next
    VTWriteWordHealth
    On Error GoTo 0
End Sub

Private Sub VTWriteWordHealth()
    Dim statusPath As String
    Dim payload As String
    statusPath = VTApplicationSupportRoot() & VT_WORD_STATUS_FILE
    payload = "{" & _
        """loaded"":true," & _
        """pluginVersion"":" & VTJsonString(VT_PLUGIN_VERSION) & "," & _
        """sourceRevision"":" & _
            VTJsonString(VT_WORD_SOURCE_REVISION) & "," & _
        """host"":""word""," & _
        """timestamp"":" & VTJsonString(Format$(Now, "yyyy-mm-dd\Thh:nn:ss")) & _
        "}"
    VTWriteTextAtomic statusPath, payload
End Sub
