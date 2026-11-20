Attribute VB_Name = "VTWordDocumentImport"
Option Explicit
Option Private Module

' Structured document import: the manifest, paragraphs and formulas, progress and cancelling.

Private Const VT_WORD_DOCUMENT_IMPORT_BOOKMARK_PREFIX As String = "VT_D_"
Private VT_WORD_V2_IMPORT_CACHE_DOCUMENT As String
Private VT_WORD_V2_IMPORT_CACHE_POSITION As Long

Public Function VTDocumentImportBookmarkName(ByVal sessionId As String) As String
    If Not VTIsCanonicalUuid(sessionId) Then
        Err.Raise vbObjectError + 7580, "VisualTeX", _
            "The document import Session id is invalid."
    End If
    VTDocumentImportBookmarkName = _
        VT_WORD_DOCUMENT_IMPORT_BOOKMARK_PREFIX & _
        Replace$(Left$(sessionId, 24), "-", "")
    If Len(VTDocumentImportBookmarkName) > 40 Then
        Err.Raise vbObjectError + 7580, "VisualTeX", _
            "The document import Bookmark name is too long."
    End If
End Function

Public Function VTDocumentImportRequestJson( _
    ByVal sessionId As String, _
    ByVal sourceDocumentId As String, _
    ByVal bookmarkName As String, _
    ByVal defaultFontSizePt As Double) As String

    VTDocumentImportRequestJson = "{" & _
        """protocolVersion"":" & CStr(VT_PROTOCOL_VERSION) & "," & _
        """sessionId"":" & VTJsonString(sessionId) & "," & _
        """host"":""word""," & _
        """mode"":""create""," & _
        """operation"":""documentImport""," & _
        """formulaId"":null," & _
        """displayMode"":""inline""," & _
        """numbered"":false," & _
        """nativeEquation"":false,"
    VTDocumentImportRequestJson = _
        VTDocumentImportRequestJson & _
        """sourceDocumentId"":" & VTJsonString(sourceDocumentId) & "," & _
        """sourceObjectId"":null," & _
        """encodedMetadata"":null," & _
        """pendingMarker"":null," & _
        """fontSizePt"":null," & _
        """referenceWidthPt"":null," & _
        """referenceHeightPt"":null," & _
        """powerPoint"":null,"
    VTDocumentImportRequestJson = _
        VTDocumentImportRequestJson & _
        """documentImport"":{" & _
        """bookmarkName"":" & VTJsonString(bookmarkName) & "," & _
        """defaultFontSizePt"":" & VTJsonNumber(defaultFontSizePt) & _
        "}}"
End Function

Public Function VTParseDocumentImportManifestText( _
    ByVal contents As String, _
    ByVal sessionId As String) As Object

    Dim dictionary As Object
    Dim rows() As String
    Dim row As Variant
    Dim separator As Long
    Dim key As String
    Dim value As String

    If Len(contents) = 0 Or Len(contents) > 16777216 Then
        Err.Raise vbObjectError + 7581, "VisualTeX", _
            "The document import manifest has an invalid size."
    End If
    Set dictionary = New Collection
    contents = Replace$(contents, vbCrLf, vbLf)
    contents = Replace$(contents, vbCr, vbLf)
    rows = Split(contents, vbLf)
    For Each row In rows
        If Len(CStr(row)) > 0 Then
            separator = InStr(1, CStr(row), "=", vbBinaryCompare)
            If separator <= 1 Then
                Err.Raise vbObjectError + 7581, "VisualTeX", _
                    "The document import manifest contains an invalid row."
            End If
            key = Left$(CStr(row), separator - 1)
            value = Mid$(CStr(row), separator + 1)
            If VTCollectionHasKey(dictionary, key) Then
                Err.Raise vbObjectError + 7581, "VisualTeX", _
                    "The document import manifest contains a duplicate key."
            End If
            dictionary.Add value, key
        End If
    Next row
    VTRequireDispatchValue dictionary, "protocolVersion"
    VTRequireDispatchValue dictionary, "sessionId"
    VTRequireDispatchValue dictionary, "operation"
    VTRequireDispatchValue dictionary, "outputKind"
    VTRequireDispatchValue dictionary, "sourceDocumentId"
    VTRequireDispatchValue dictionary, "bookmarkName"
    VTRequireDispatchValue dictionary, "itemCount"
    If CStr(dictionary("protocolVersion")) <> _
       CStr(VT_PROTOCOL_VERSION) Or _
       CStr(dictionary("sessionId")) <> sessionId Then
        Err.Raise vbObjectError + 7581, "VisualTeX", _
            "The document import manifest identity is invalid."
    End If
    Set VTParseDocumentImportManifestText = dictionary
End Function

Public Function VTReadDocumentImportManifest( _
    ByVal manifestPath As String, _
    ByVal sessionId As String) As Object

    VTValidateAbsoluteVisualTeXPath manifestPath
    Set VTReadDocumentImportManifest = _
        VTParseDocumentImportManifestText( _
            VTReadText(manifestPath, 16777216), sessionId)
End Function

Public Function VTDocumentImportItemKey( _
    ByVal itemIndex As Long, _
    ByVal suffix As String) As String

    If itemIndex < 0 Or Len(suffix) = 0 Then
        Err.Raise vbObjectError + 7582, "VisualTeX", _
            "The document import item key is invalid."
    End If
    VTDocumentImportItemKey = "item" & CStr(itemIndex) & suffix
End Function

Public Function VTDocumentImportRequired( _
    ByVal manifest As Object, _
    ByVal itemIndex As Long, _
    ByVal suffix As String) As String

    Dim key As String
    key = VTDocumentImportItemKey(itemIndex, suffix)
    VTRequireDispatchValue manifest, key
    VTDocumentImportRequired = CStr(manifest(key))
End Function

Private Sub VTDocumentImportInsertText( _
    ByRef cursorRange As Range, _
    ByVal textBase64 As String)

    Dim plainText As String
    Dim insertionRange As Range
    Dim insertedRange As Range
    Dim insertionStart As Long
    Dim replacesInlineMathAnchor As Boolean

    If cursorRange Is Nothing Then
        Err.Raise vbObjectError + 7583, "VisualTeX", _
            "The document import text insertion point is missing."
    End If
    plainText = VTBase64UrlDecodeUtf8(textBase64)
    If Len(plainText) = 0 Then Exit Sub
    plainText = Replace$(plainText, vbCrLf, vbLf)
    plainText = Replace$(plainText, vbCr, vbLf)
    ' One manifest text item represents one logical Word paragraph. Preserve
    ' any internal source line breaks as manual line breaks instead of creating
    ' extra Word paragraphs behind the paragraph metadata's back. Real paragraph
    ' boundaries are created only by VTFinalizeDocumentImportParagraph.
    plainText = Replace$(plainText, vbLf, Chr$(11))

    insertionStart = cursorRange.Start
    replacesInlineMathAnchor = _
        cursorRange.End = cursorRange.Start + 1 And _
        cursorRange.Text = ChrW(8288) And _
        cursorRange.OMaths.Count = 0
    If replacesInlineMathAnchor Then
        ' VTPlaceCaretAfterInlineNativeEquation selects a U+2060 boundary
        ' outside OMath. Replace that temporary marker with the next prose
        ' fragment so the paragraph contains no hidden separator character.
        Set insertionRange = cursorRange.Duplicate
        insertionRange.Text = plainText
    Else
        Set insertionRange = VTNativeInsertionBoundary(cursorRange)
        insertionStart = insertionRange.Start
        insertionRange.Text = plainText
    End If
    Set insertedRange = cursorRange.Document.Range( _
        Start:=insertionStart, End:=insertionStart + Len(plainText))
    ' Imported prose is plain Word text. It must not inherit an italic, bold,
    ' raised-baseline, or underline state left behind by the insertion caret or
    ' by an adjacent native equation.
    On Error Resume Next
    insertedRange.Font.Reset
    insertedRange.Font.Bold = False
    insertedRange.Font.Italic = False
    insertedRange.Font.Underline = wdUnderlineNone
    insertedRange.Font.Position = 0
    On Error GoTo 0
    Set cursorRange = cursorRange.Document.Range( _
        Start:=insertionStart + Len(plainText), _
        End:=insertionStart + Len(plainText))
End Sub

Private Function VTDocumentImportParagraphValue( _
    ByVal manifest As Object, _
    ByVal itemIndex As Long, _
    ByVal suffix As String) As String

    VTDocumentImportParagraphValue = VTDispatchOptional( _
        manifest, VTDocumentImportItemKey(itemIndex, suffix))
End Function

Private Function VTDocumentImportParagraphFlag( _
    ByVal manifest As Object, _
    ByVal itemIndex As Long, _
    ByVal suffix As String) As Boolean

    Dim value As String
    value = VTDocumentImportParagraphValue(manifest, itemIndex, suffix)
    If Len(value) = 0 Or value = "0" Then
        VTDocumentImportParagraphFlag = False
    ElseIf value = "1" Then
        VTDocumentImportParagraphFlag = True
    Else
        Err.Raise vbObjectError + 7587, "VisualTeX", _
            "A document paragraph Boolean value is invalid."
    End If
End Function

Private Sub VTApplyDocumentImportParagraphFormat( _
    ByVal paragraphRange As Range, _
    ByVal paragraphStyle As String, _
    ByVal paragraphAlignment As String, _
    ByVal listKind As String, _
    ByVal listLevel As Long)

    Dim levelIndex As Long

    If paragraphRange Is Nothing Then
        Err.Raise vbObjectError + 7587, "VisualTeX", _
            "The imported Word paragraph Range is missing."
    End If
    Select Case paragraphStyle
        Case "", "normal"
            paragraphRange.Style = wdStyleNormal
        Case "heading1"
            paragraphRange.Style = wdStyleHeading1
        Case "heading2"
            paragraphRange.Style = wdStyleHeading2
        Case "heading3"
            paragraphRange.Style = wdStyleHeading3
        Case "heading4"
            paragraphRange.Style = wdStyleHeading4
        Case "quote"
            paragraphRange.Style = wdStyleNormal
            paragraphRange.ParagraphFormat.LeftIndent = 36!
            paragraphRange.ParagraphFormat.RightIndent = 36!
        Case "code"
            paragraphRange.Style = wdStyleNormal
            paragraphRange.Font.Name = "Menlo"
        Case Else
            Err.Raise vbObjectError + 7587, "VisualTeX", _
                "The imported Word paragraph style is invalid."
    End Select

    Select Case paragraphAlignment
        Case "", "left"
            paragraphRange.ParagraphFormat.Alignment = wdAlignParagraphLeft
        Case "center"
            paragraphRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Case "right"
            paragraphRange.ParagraphFormat.Alignment = wdAlignParagraphRight
        Case "justify"
            paragraphRange.ParagraphFormat.Alignment = wdAlignParagraphJustify
        Case Else
            Err.Raise vbObjectError + 7587, "VisualTeX", _
                "The imported Word paragraph alignment is invalid."
    End Select

    Select Case listKind
        Case "", "none"
            On Error Resume Next
            paragraphRange.ListFormat.RemoveNumbers
            On Error GoTo 0
        Case "bullet"
            paragraphRange.ListFormat.ApplyBulletDefault
        Case "number"
            paragraphRange.ListFormat.ApplyNumberDefault
        Case Else
            Err.Raise vbObjectError + 7587, "VisualTeX", _
                "The imported Word list kind is invalid."
    End Select
    If listKind <> "" And listKind <> "none" Then
        If listLevel < 1 Or listLevel > 9 Then
            Err.Raise vbObjectError + 7587, "VisualTeX", _
                "The imported Word list level is invalid."
        End If
        For levelIndex = 2 To listLevel
            paragraphRange.ListFormat.ListIndent
        Next levelIndex
    End If

    If paragraphStyle = "normal" Or paragraphStyle = "quote" Or _
       paragraphStyle = "code" Or Len(paragraphStyle) = 0 Then
        paragraphRange.ParagraphFormat.SpaceBefore = 0!
        paragraphRange.ParagraphFormat.SpaceAfter = 0!
    End If
End Sub

Private Sub VTPrepareDocumentImportParagraph( _
    ByVal cursorRange As Range, _
    ByVal paragraphStyle As String, _
    ByVal paragraphAlignment As String, _
    ByVal listKind As String, _
    ByVal listLevel As Long)

    Dim paragraphRange As Range

    If cursorRange Is Nothing Then
        Err.Raise vbObjectError + 7587, "VisualTeX", _
            "The document paragraph insertion point is missing."
    End If
    Set paragraphRange = cursorRange.Paragraphs(1).Range.Duplicate
    ' Word for Mac can reject ListFormat.ApplyBulletDefault/ApplyNumberDefault
    ' with error 5904 when the Range already contains an OMath object. Apply
    ' the paragraph style and list template while the target paragraph is still
    ' empty, then insert prose and independent inline formulas into it.
    VTApplyDocumentImportParagraphFormat _
        paragraphRange, paragraphStyle, paragraphAlignment, listKind, listLevel
End Sub

Private Sub VTFinalizeDocumentImportParagraph( _
    ByRef cursorRange As Range, _
    ByVal paragraphStart As Long, _
    ByVal paragraphStyle As String, _
    ByVal paragraphAlignment As String, _
    ByVal listKind As String, _
    ByVal listLevel As Long, _
    ByVal sessionId As String, _
    ByVal itemIndex As Long)

    Dim documentObject As Document
    Dim contentRange As Range
    Dim paragraphRange As Range
    Dim boundaryRange As Range
    Dim nextStart As Long

    If cursorRange Is Nothing Then
        Err.Raise vbObjectError + 7587, "VisualTeX", _
            "The document paragraph insertion point is missing."
    End If
    VTTraceDocumentImportStage sessionId, "paragraph-range-start", itemIndex
    Set documentObject = cursorRange.Document
    If cursorRange.End = cursorRange.Start + 1 And _
       cursorRange.Text = ChrW(8288) And _
       cursorRange.OMaths.Count = 0 Then
        ' A paragraph ending with an inline equation still owns the verified
        ' U+2060 boundary selected by VTPlaceCaretAfterInlineNativeEquation.
        ' Replace it with the real paragraph mark instead of leaving a hidden
        ' character or collapsing the insertion point back into OMath.
        nextStart = cursorRange.Start
        cursorRange.Text = vbCr
        Set cursorRange = documentObject.Range( _
            Start:=nextStart, End:=nextStart)
    End If
    If paragraphStart < 0 Or paragraphStart > cursorRange.Start Then
        Err.Raise vbObjectError + 7587, "VisualTeX", _
            "The document paragraph boundary is invalid."
    End If
    Set contentRange = documentObject.Range( _
        Start:=paragraphStart, End:=cursorRange.Start)
    VTTraceDocumentImportStage sessionId, "paragraph-range-created", itemIndex
    If contentRange.Paragraphs.Count <> 1 Then
        Err.Raise vbObjectError + 7587, "VisualTeX", _
            "An imported text paragraph crossed an unexpected Word paragraph boundary."
    End If
    Set paragraphRange = contentRange.Paragraphs(1).Range.Duplicate
    ' Paragraph/list formatting was applied before any content was inserted.
    ' Reapplying a list template here would make a Range containing native
    ' equations non-editable in Word for Mac (runtime error 5904).
    VTTraceDocumentImportStage sessionId, "paragraph-format-preserved", itemIndex

    nextStart = cursorRange.Start
    If nextStart < documentObject.Content.End Then
        Set boundaryRange = documentObject.Range( _
            Start:=nextStart, End:=nextStart + 1)
        If boundaryRange.Text = vbCr Then
            If boundaryRange.End < documentObject.Content.End Then
                nextStart = boundaryRange.End
            Else
                ' Word's final paragraph mark cannot be used as a collapsed
                ' Range at Content.End. Insert one real paragraph boundary
                ' before it, then keep the caret before the final mark.
                Set boundaryRange = documentObject.Range( _
                    Start:=nextStart, End:=nextStart)
                boundaryRange.InsertAfter vbCr
                nextStart = nextStart + 1
            End If
        Else
            Set boundaryRange = documentObject.Range( _
                Start:=nextStart, End:=nextStart)
            boundaryRange.InsertAfter vbCr
            nextStart = nextStart + 1
        End If
    Else
        Set boundaryRange = documentObject.Range( _
            Start:=documentObject.Content.End - 1, _
            End:=documentObject.Content.End - 1)
        boundaryRange.InsertAfter vbCr
        nextStart = documentObject.Content.End - 1
    End If
    If nextStart >= documentObject.Content.End Then
        nextStart = documentObject.Content.End - 1
    End If
    VTTraceDocumentImportStage sessionId, "paragraph-boundary-complete", itemIndex
    Set cursorRange = documentObject.Range( _
        Start:=nextStart, End:=nextStart)
    VTTraceDocumentImportStage sessionId, "paragraph-cursor-complete", itemIndex
End Sub

Public Sub VTDeleteDocumentImportedFormulaState( _
    ByVal documentObject As Document, _
    ByVal formulaId As String)

    Dim bookmarkName As String

    On Error Resume Next
    VTDeleteWordLatexPayload documentObject, formulaId
    VTDeleteWordOmmlPayload documentObject, formulaId
    VTDeleteWordMetadataPayload documentObject, formulaId
    VTDeleteDocumentVariable documentObject, VTWordFormatVariableName(formulaId)
    VTDeleteDocumentVariable documentObject, VTWordImageScaleVariableName(formulaId)
    VTDeleteDocumentVariable documentObject, VTWordImageCanvasVariableName(formulaId)
    VTDeleteDocumentVariable _
        documentObject, VTWordImageInkCenterVariableName(formulaId)
    VTDeleteDocumentVariable _
        documentObject, VTWordImageLetterFontVariableName(formulaId)
    VTDeleteDocumentVariable _
        documentObject, VTWordNativeSignatureVariableName(formulaId)
    bookmarkName = VTNativeFormulaBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        documentObject.Bookmarks(bookmarkName).Delete
    End If
    bookmarkName = VTEquationCaptionBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        documentObject.Bookmarks(bookmarkName).Delete
    End If
    On Error GoTo 0
End Sub

Private Sub VTDocumentImportInsertFormula( _
    ByRef cursorRange As Range, _
    ByVal manifest As Object, _
    ByVal itemIndex As Long, _
    ByVal outputKind As String, _
    ByRef insertedFormulaIds As Collection, _
    Optional ByVal overrideFontSizePt As Double = 0#, _
    Optional ByVal preserveParagraphTopology As Boolean = False, _
    Optional ByVal sharedVectorDocument As Variant, _
    Optional ByVal sharedNativeDocument As Variant, _
    Optional ByVal deferNumberingReconcile As Boolean = False)

    Dim targetDocument As Document
    Dim targetRange As Range
    Dim formulaRange As Range
    Dim candidate As InlineShape
    Dim numberLayoutRange As Range
    Dim formulaId As String
    Dim latexBase64 As String
    Dim ommlBase64 As String
    Dim nativeDocumentPath As String
    Dim nativeBatchDocumentIndex As Long
    Dim displayMode As String
    Dim numbered As Boolean
    Dim metadata As String
    Dim imagePath As String
    Dim v2SequenceName As String
    Dim v2HeadingLevel As Long
    Dim v2Separator As String
    Dim v2LiteralPrefix As String
    Dim v2Automatic As Boolean
    Dim v2Restart As Long
    Dim v2Initial As Long
    Dim v2TransferEquation As Range
    Dim vectorDocumentPath As String
    Dim vectorDocumentIndex As Long
    Dim fallbackImagePath As String
    Dim captionText As String
    Dim formulaReference As String
    Dim fontSizePt As Double
    Dim widthPoints As Double
    Dim heightPoints As Double
    Dim baselinePoints As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim inkCenterYRatio As Double
    Dim observedWordFontSizePt As Double
    Dim fontScale As Double
    Dim candidateStart As Long
    Dim numberCreated As Boolean
    Dim nativeMath As OMath
    Dim formulaStage As String
    Dim formulaErrorNumber As Long
    Dim formulaErrorDescription As String
    Dim joinedNative As Boolean

    On Error GoTo FormulaFailed
    formulaStage = "validate-input"
    If cursorRange Is Nothing Then
        Err.Raise vbObjectError + 7584, "VisualTeX", _
            "The document formula insertion point is missing."
    End If
    Set targetDocument = cursorRange.Document
    formulaId = VTDocumentImportRequired( _
        manifest, itemIndex, "formulaId")
    latexBase64 = VTDocumentImportRequired( _
        manifest, itemIndex, "latexBase64")
    displayMode = VTDocumentImportRequired( _
        manifest, itemIndex, "displayMode")
    numbered = (VTDocumentImportRequired( _
        manifest, itemIndex, "numbered") = "1")
    fontSizePt = VTDispatchPositiveDouble( _
        manifest, VTDocumentImportItemKey(itemIndex, "fontSizePt"))
    metadata = VTDocumentImportRequired( _
        manifest, itemIndex, "metadata")
    ommlBase64 = VTDocumentImportRequired( _
        manifest, itemIndex, "ommlBase64")
    nativeDocumentPath = VTDocumentImportRequired( _
        manifest, itemIndex, "nativeDocumentPath")
    widthPoints = VTDispatchPositiveDouble( _
        manifest, VTDocumentImportItemKey(itemIndex, "widthPoints"))
    heightPoints = VTDispatchPositiveDouble( _
        manifest, VTDocumentImportItemKey(itemIndex, "heightPoints"))
    baselinePoints = VTDispatchOptionalDouble( _
        manifest, VTDocumentImportItemKey(itemIndex, "baseline"), 0#)
    referenceWidthPt = VTDispatchPositiveDouble( _
        manifest, VTDocumentImportItemKey(itemIndex, "referenceWidthPt"))
    referenceHeightPt = VTDispatchPositiveDouble( _
        manifest, VTDocumentImportItemKey(itemIndex, "referenceHeightPt"))
    referenceBaselinePt = VTDispatchOptionalDouble( _
        manifest, _
        VTDocumentImportItemKey(itemIndex, "referenceBaselinePt"), 0#)
    inkCenterYRatio = VTDispatchOptionalDouble( _
        manifest, _
        VTDocumentImportItemKey(itemIndex, "inkCenterYRatio"), 0.5)

    If Not VTIsCanonicalUuid(formulaId) Or _
       Not VTIsBase64UrlPayload(latexBase64) Or _
       Not VTIsBase64UrlPayload(ommlBase64) Or _
       Not VTIsEncodedMetadata(metadata) Then
        Err.Raise vbObjectError + 7584, "VisualTeX", _
            "A document formula contains invalid VisualTeX metadata."
    End If
    If displayMode <> "inline" And displayMode <> "block" Then
        Err.Raise vbObjectError + 7584, "VisualTeX", _
            "A document formula has an invalid display mode."
    End If
    If numbered And displayMode <> "block" Then
        Err.Raise vbObjectError + 7584, "VisualTeX", _
            "Only document display formulas can be numbered."
    End If
    If Not VTValidWordFormulaFontSize(fontSizePt) Or _
       referenceWidthPt <= 0# Or referenceHeightPt <= 0# Or _
       referenceBaselinePt < -256# Or referenceBaselinePt > 0# Or _
       inkCenterYRatio < 0# Or inkCenterYRatio > 1# Then
        Err.Raise vbObjectError + 7584, "VisualTeX", _
            "A document formula has invalid point-size geometry."
    End If
    If overrideFontSizePt <> 0# Then
        If Not VTValidWordFormulaFontSize(overrideFontSizePt) Then
            Err.Raise vbObjectError + 7584, "VisualTeX", _
                "The redraw source formula has an invalid Word font size."
        End If
        fontScale = overrideFontSizePt / fontSizePt
        widthPoints = widthPoints * fontScale
        heightPoints = heightPoints * fontScale
        ' Recompute from the fractional 14 pt reference instead of scaling an
        ' already-rounded Word Position value. This keeps redraw/import baseline
        ' rounding identical to a freshly rendered formula at the new size.
        baselinePoints = referenceBaselinePt * overrideFontSizePt / _
            VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
        fontSizePt = overrideFontSizePt
    End If
    VTValidateAbsoluteVisualTeXPath nativeDocumentPath
    If Not VTPathFileExists(nativeDocumentPath) Then
        Err.Raise vbObjectError + 7584, "VisualTeX", _
            "A document formula native Word staging file is missing."
    End If

    formulaReference = VTFormulaReference( _
        formulaId, displayMode, numbered)
    captionText = VTEquationCrossReferenceText(latexBase64)
    Set targetRange = cursorRange.Duplicate
    targetRange.Collapse wdCollapseStart
    If displayMode = "block" Then
        If preserveParagraphTopology Then
            ' A format-only range replacement already owns the source formula's
            ' dedicated paragraph. Keep the exact insertion point in that same
            ' paragraph so an existing external VT_N_ SEQ helper remains
            ' immediately adjacent to its visible numbered formula row.
        ElseIf Not numbered Then
            Set targetRange = _
                VTPrepareBatchUnnumberedDisplayInsertionRange(targetRange)
        Else
            Set targetRange = VTPrepareWordCreateInsertionRange( _
                targetRange, displayMode)
        End If
    End If
    If outputKind = "omml" And numbered And displayMode = "block" Then
        ' Equation format v2: the app builds the complete numbered equation
        ' (as for an editor commit); it is inserted from its own package.
        formulaStage = "v2-plan-number"
        VTV2PlanInsertion targetRange.Document, targetRange.Start, v2SequenceName, _
            v2HeadingLevel, v2Separator, v2LiteralPrefix, v2Automatic, v2Restart, _
            v2Initial, v2TransferEquation
        formulaStage = "v2-build-number"
        nativeDocumentPath = VTV2RequestNumberDocx(CStr(manifest("sessionId")), _
            targetRange.Document, ommlBase64, 1, VTV2NumberItemManifest(1, _
                v2SequenceName, v2HeadingLevel, v2Separator, v2LiteralPrefix, _
                v2Automatic, v2Restart, v2Initial))
    End If
    If outputKind = "omml" Then
        formulaStage = "insert-native"
        If IsObject(sharedNativeDocument) And Not (numbered And displayMode = "block") Then
            nativeBatchDocumentIndex = CLng(Val(VTDispatchOptional( _
                manifest, _
                VTDocumentImportItemKey(itemIndex, "nativeBatchDocumentIndex"))))
            Set formulaRange = VTInsertNativeEquationFromSharedDocument( _
                targetRange, sharedNativeDocument, nativeBatchDocumentIndex, _
                displayMode, displayMode = "block", joinedNative, fontSizePt)
        Else
            Set formulaRange = VTInsertNativeEquationAtRange( _
                targetRange, ommlBase64, nativeDocumentPath, _
                displayMode, displayMode = "block", False, _
                joinedNative:=joinedNative, nativeFontSizePt:=fontSizePt)
        End If
        ' Word for Mac can invalidate a live OMath Range while document
        ' Variables are added or replaced. Persist the formula identity before
        ' writing state, then reacquire the OMath from the durable Bookmark.
        formulaStage = "bookmark-native-before-state"
        VTSetNativeFormulaBookmark targetDocument, formulaRange, formulaId
        ' The live Word tree is the native equation's only persistent source.
        ' This bookmark is a transaction anchor, removed before acknowledgement.
        formulaStage = "resolve-native-after-state"
        Set nativeMath = VTNativeMathForBookmark( _
            targetDocument.Bookmarks( _
                VTNativeFormulaBookmarkName(formulaId)))
        If nativeMath Is Nothing Then
            Err.Raise vbObjectError + 7584, "VisualTeX", _
                "Word lost the imported native equation after storing its state."
        End If
        Set formulaRange = nativeMath.Range.Duplicate
        ' VTInsertNativeEquationAtRange already applies the requested inline or
        ' display type. Only numbered display equations need extra layout here.
        If displayMode = "block" And numbered Then
            formulaStage = "number-native"
            VTV2FinishInsertedNumber formulaRange.Document, formulaRange, _
                v2TransferEquation
            Set nativeMath = VTNativeMathForBookmark( _
                targetDocument.Bookmarks( _
                    VTNativeFormulaBookmarkName(formulaId)))
            If nativeMath Is Nothing Then
                Err.Raise vbObjectError + 7584, "VisualTeX", _
                    "Word lost the numbered imported native equation."
            End If
            Set formulaRange = nativeMath.Range.Duplicate
        End If
        formulaStage = "apply-native-font"
        If Not joinedNative Then VTV2ApplyEquationTypography formulaRange, fontSizePt
        formulaStage = "bookmark-native-final"
        VTSetNativeFormulaBookmark targetDocument, formulaRange, formulaId
        VTDeleteWordImageOwnerBookmark targetDocument, formulaId
        formulaStage = "resolve-native-before-signature"
        Set nativeMath = VTNativeMathForBookmark( _
            targetDocument.Bookmarks( _
                VTNativeFormulaBookmarkName(formulaId)))
        If nativeMath Is Nothing Then
            Err.Raise vbObjectError + 7584, "VisualTeX", _
                "Word lost the imported native equation before structure caching."
        End If
        formulaStage = "write-native-signature"
        Set formulaRange = nativeMath.Range.Duplicate
        insertedFormulaIds.Add formulaId
        formulaStage = "place-native-caret"
        If displayMode = "block" Then
            If preserveParagraphTopology Then
                Set cursorRange = targetDocument.Range( _
                    Start:=formulaRange.End, End:=formulaRange.End)
            ElseIf numbered Then
                Set cursorRange = VTContinuationRangeAfterDisplayFormula( _
                    formulaRange, formulaId)
            Else
                Set cursorRange = _
                    VTBatchContinuationRangeAfterUnnumberedDisplayFormula( _
                        formulaRange)
            End If
        Else
            ' A collapsed Range at OMath.Range.End remains inside the equation
            ' on Word for Mac. Establish a verified text boundary so a later
            ' inline formula in the same paragraph creates a separate OMath.
            VTPlaceCaretAfterInlineNativeEquation formulaRange
            Set cursorRange = Selection.Range.Duplicate
        End If
        VTFinishPureNativeEquation formulaRange, formulaId
        Exit Sub
    End If

    If outputKind <> "image" Then
        Err.Raise vbObjectError + 7584, "VisualTeX", _
            "The document formula output kind is invalid."
    End If
    imagePath = VTDocumentImportRequired( _
        manifest, itemIndex, "imagePath")
    vectorDocumentPath = VTDocumentImportRequired( _
        manifest, itemIndex, "vectorDocumentPath")
    vectorDocumentIndex = CLng(Val(VTDispatchOptional( _
        manifest, VTDocumentImportItemKey(itemIndex, "vectorDocumentIndex"))))
    fallbackImagePath = VTDocumentImportRequired( _
        manifest, itemIndex, "fallbackImagePath")
    VTValidateAbsoluteVisualTeXPath imagePath
    VTValidateAbsoluteVisualTeXPath vectorDocumentPath
    VTValidateAbsoluteVisualTeXPath fallbackImagePath
    If Not VTPathFileExists(imagePath) Or _
       Not VTPathFileExists(vectorDocumentPath) Or _
       Not VTPathFileExists(fallbackImagePath) Then
        Err.Raise vbObjectError + 7584, "VisualTeX", _
            "A document formula SVG staging file is missing."
    End If

    Set candidate = VTAddWordFormulaPicture( _
        targetDocument, targetRange, vectorDocumentPath, _
        fallbackImagePath, sharedVectorDocument, vectorDocumentIndex)
    candidateStart = candidate.Range.Start
    candidate.LockAspectRatio = msoFalse
    candidate.Width = CSng(widthPoints)
    candidate.Height = CSng(heightPoints)
    candidate.LockAspectRatio = msoTrue
    candidate.AlternativeText = metadata
    candidate.Title = formulaReference
    On Error Resume Next
    candidate.Range.Font.Size = CSng(fontSizePt)
    On Error GoTo FormulaFailed
    If displayMode = "inline" Then
        VTApplyWordInlineImageBaseline _
            candidate, referenceHeightPt, referenceBaselinePt, fontSizePt
    Else
        candidate.Range.Font.Position = 0
        candidate.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If

    VTSetWordLatexPayload targetDocument, formulaId, latexBase64
    VTSetWordOmmlPayload targetDocument, formulaId, ommlBase64
    VTSetWordMetadataPayload targetDocument, formulaId, metadata
    VTSetWordFormulaFormat _
        targetDocument, formulaId, displayMode, numbered
    observedWordFontSizePt = VTInlineShapeWordFontSize(candidate)
    If Not VTValidWordFormulaFontSize(observedWordFontSizePt) Then
        observedWordFontSizePt = fontSizePt
    End If
    VTSetWordImageScaleState _
        targetDocument, formulaId, fontSizePt, _
        referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
        observedWordFontSizePt
    VTSetWordImageInkCenterYRatio _
        targetDocument, formulaId, inkCenterYRatio
    VTSetWordImageLetterFont targetDocument, formulaId, VTDispatchOptional( _
        manifest, VTDocumentImportItemKey(itemIndex, "formulaLetterFont"))

    ' Document.Variable updates can invalidate an InlineShape COM wrapper on
    ' Word for Mac even though the drawing remains in the document. Resolve the
    ' newly committed image from its durable metadata before field wrapping.
    Set targetRange = targetDocument.Range( _
        Start:=candidateStart, End:=candidateStart + 1)
    If targetRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7584, "VisualTeX", _
            "Word could not re-resolve the imported formula image before MacroButton wrapping."
    End If
    Set candidate = targetRange.InlineShapes(1)
    If candidate.AlternativeText <> metadata Or _
       candidate.Title <> formulaReference Then
        Err.Raise vbObjectError + 7584, "VisualTeX", _
            "Word re-resolved the wrong imported formula image."
    End If
    Set candidate = VTEnsureVisualTeXImageMacroButton(candidate)
    If displayMode = "inline" Then
        VTApplyWordInlineImageBaseline _
            candidate, referenceHeightPt, referenceBaselinePt, fontSizePt
    End If

    If displayMode = "block" Then
        If numbered Then
            numberCreated = False
            Set numberLayoutRange = VTEnsureImageEquationNumber( _
                candidate, heightPoints, formulaId, captionText, _
                numberCreated, deferNumberingReconcile)
        Else
            VTNormalizeUnnumberedDisplayParagraph candidate.Range
        End If
        VTNormalizeImageDisplayParagraph candidate.Range
    End If
    VTSetWordImageOwnerBookmark targetDocument, candidate, formulaId
    insertedFormulaIds.Add formulaId

    If displayMode = "block" Then
        If preserveParagraphTopology Then
            Set formulaRange = VTVisualTeXImageContainerRange(candidate)
            If formulaRange Is Nothing Then Set formulaRange = candidate.Range.Duplicate
            Set cursorRange = targetDocument.Range( _
                Start:=formulaRange.End, End:=formulaRange.End)
        ElseIf numbered Then
            Set cursorRange = VTContinuationRangeAfterDisplayFormula( _
                candidate.Range, formulaId)
        Else
            Set cursorRange = _
                VTBatchContinuationRangeAfterUnnumberedDisplayFormula( _
                    candidate.Range)
        End If
    Else
        Set formulaRange = VTVisualTeXImageContainerRange(candidate)
        Set cursorRange = targetDocument.Range( _
            Start:=formulaRange.End, End:=formulaRange.End)
    End If
    Exit Sub

FormulaFailed:
    formulaErrorNumber = Err.Number
    formulaErrorDescription = Err.Description
    On Error GoTo 0
    Err.Raise formulaErrorNumber, _
        "VisualTeX Word document formula import", _
        formulaStage & ": " & formulaErrorDescription
End Sub

Private Sub VTCleanupFailedDocumentImportFormula( _
    ByVal documentObject As Document, _
    ByVal formulaId As String)

    Dim nativeBookmarkName As String
    Dim nativeRange As Range
    Dim imageRange As Range
    Dim formulaShape As InlineShape
    Dim candidateFormulaId As String
    Dim candidateDisplayMode As String
    Dim candidateNumbered As Boolean

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then Exit Sub
    On Error Resume Next
    nativeBookmarkName = VTNativeFormulaBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(nativeBookmarkName) Then
        Set nativeRange = documentObject.Bookmarks( _
            nativeBookmarkName).Range.Duplicate
    End If
    For Each formulaShape In documentObject.InlineShapes
        candidateFormulaId = ""
        candidateDisplayMode = ""
        candidateNumbered = False
        If VTTryParseFormulaReference( _
           formulaShape.Title, candidateFormulaId, _
           candidateDisplayMode, candidateNumbered) Then
            If StrComp(candidateFormulaId, formulaId, vbTextCompare) = 0 Then
                Set imageRange = VTVisualTeXImageContainerRange(formulaShape)
                Exit For
            End If
        End If
    Next formulaShape
    If Not imageRange Is Nothing Then imageRange.Delete
    If Not nativeRange Is Nothing Then nativeRange.Delete
    VTDeleteEquationNumberScaffold documentObject, formulaId, False
    VTDeleteDocumentImportedFormulaState documentObject, formulaId
    Err.Clear
    On Error GoTo 0
End Sub

Public Function VTTryDocumentImportInsertFormula( _
    ByRef cursorRange As Range, _
    ByVal manifest As Object, _
    ByVal itemIndex As Long, _
    ByVal outputKind As String, _
    ByRef insertedFormulaIds As Collection, _
    ByRef failureDescription As String, _
    Optional ByVal overrideFontSizePt As Double = 0#, _
    Optional ByVal preserveParagraphTopology As Boolean = False, _
    Optional ByVal sharedVectorDocument As Variant, _
    Optional ByVal sharedNativeDocument As Variant, _
    Optional ByVal deferCleanupToOuterUndo As Boolean = False, _
    Optional ByVal deferNumberingReconcile As Boolean = False) As Boolean

    Dim documentObject As Document
    Dim formulaId As String
    Dim insertionStart As Long
    Dim failureNumber As Long
    Dim nativeBackup As Document
    Dim sourceParagraph As Range
    Dim backupParagraph As Range
    Dim sourceText As String
    Dim sourceLength As Long

    failureDescription = ""
    If cursorRange Is Nothing Then Exit Function
    Set documentObject = cursorRange.Document
    insertionStart = cursorRange.Start
    formulaId = VTDocumentImportRequired(manifest, itemIndex, "formulaId")
    On Error GoTo Failed
    If outputKind = "omml" And _
       VTDocumentImportRequired(manifest, itemIndex, "displayMode") = "inline" Then
        If Not VTAdjacentNativeMath(cursorRange) Is Nothing Then
            Set sourceParagraph = cursorRange.Paragraphs(1).Range.Duplicate
            sourceText = sourceParagraph.Text
            sourceLength = sourceParagraph.End - sourceParagraph.Start
            Set nativeBackup = Documents.Add(Visible:=False)
            Set backupParagraph = nativeBackup.Content.Duplicate
            backupParagraph.Collapse wdCollapseStart
            backupParagraph.FormattedText = sourceParagraph.FormattedText
            Set backupParagraph = nativeBackup.Range(0, sourceLength)
            If backupParagraph.Text <> sourceText Then _
                Err.Raise vbObjectError + 7472, , "Word could not back up the native insertion paragraph."
            documentObject.Activate
        End If
    End If
    VTDocumentImportInsertFormula _
        cursorRange, manifest, itemIndex, outputKind, insertedFormulaIds, _
        overrideFontSizePt, preserveParagraphTopology, _
        sharedVectorDocument, sharedNativeDocument, _
        deferNumberingReconcile
    VTTryDocumentImportInsertFormula = True
    If Not nativeBackup Is Nothing Then nativeBackup.Close SaveChanges:=wdDoNotSaveChanges
    Exit Function

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Not nativeBackup Is Nothing Then
        sourceParagraph.FormattedText = backupParagraph.FormattedText
        If sourceParagraph.Text <> sourceText Then _
            failureDescription = failureDescription & "; native paragraph rollback could not be verified"
        VTDeleteDocumentImportedFormulaState documentObject, formulaId
        nativeBackup.Close SaveChanges:=wdDoNotSaveChanges
        documentObject.Activate
    ElseIf Not deferCleanupToOuterUndo Then
        VTCleanupFailedDocumentImportFormula documentObject, formulaId
    End If
    Set cursorRange = documentObject.Range( _
        Start:=insertionStart, End:=insertionStart)
    Err.Clear
    On Error GoTo 0
    If failureNumber = 0 Then failureNumber = vbObjectError + 7584
End Function

Private Sub VTTraceDocumentImportStage( _
    ByVal sessionId As String, _
    ByVal stageName As String, _
    ByVal itemIndex As Long)

    If Not VT_WORD_TRACE_ENABLED Then Exit Sub
    On Error Resume Next
    VTWriteTextAtomic _
        VTSessionDirectory(sessionId) & "/document-import-stage.txt", _
        "stage=" & stageName & vbLf & _
        "itemIndex=" & CStr(itemIndex) & vbLf & _
        "time=" & CStr(Now) & vbLf
    Err.Clear
    On Error GoTo 0
End Sub

Private Function VTUnicodeBmp(ByVal codePoint As Long) As String
    If codePoint > &H7FFF& Then codePoint = codePoint - &H10000
    VTUnicodeBmp = ChrW$(codePoint)
End Function

Public Function VTDocumentImportStatusPrefix( _
    ByVal stageName As String) As String

    Select Case stageName
        Case "preparing"
            VTDocumentImportStatusPrefix = _
                "VisualTeX " & _
                VTUnicodeBmp(&H6B63&) & VTUnicodeBmp(&H5728&) & _
                VTUnicodeBmp(&H51C6&) & VTUnicodeBmp(&H5907&) & _
                VTUnicodeBmp(&H6279&) & VTUnicodeBmp(&H91CF&) & _
                VTUnicodeBmp(&H5BFC&) & VTUnicodeBmp(&H5165&) & _
                VTUnicodeBmp(&H2026&)
        Case "inserting"
            VTDocumentImportStatusPrefix = _
                "VisualTeX " & _
                VTUnicodeBmp(&H6B63&) & VTUnicodeBmp(&H5728&) & _
                VTUnicodeBmp(&H63D2&) & VTUnicodeBmp(&H5165&) & _
                " Word " & _
                VTUnicodeBmp(&H6587&) & VTUnicodeBmp(&H6863&) & _
                VTUnicodeBmp(&HFF1A&)
        Case "complete"
            VTDocumentImportStatusPrefix = _
                "VisualTeX " & _
                VTUnicodeBmp(&H6279&) & VTUnicodeBmp(&H91CF&) & _
                VTUnicodeBmp(&H5BFC&) & VTUnicodeBmp(&H5165&) & _
                VTUnicodeBmp(&H5B8C&) & VTUnicodeBmp(&H6210&) & _
                VTUnicodeBmp(&HFF1A&)
        Case Else
            VTDocumentImportStatusPrefix = ""
    End Select
End Function

Public Sub VTUpdateDocumentImportProgress( _
    ByVal sessionId As String, _
    ByVal currentItem As Long, _
    ByVal totalItems As Long, _
    ByVal stageName As String)

    Dim statusText As String

    If currentItem < 0 Then currentItem = 0
    If totalItems < 0 Then totalItems = 0
    If currentItem > totalItems Then currentItem = totalItems

    On Error Resume Next
    VTWriteTextAtomic _
        VTSessionDirectory(sessionId) & "/document-import-progress.txt", _
        "current=" & CStr(currentItem) & vbLf & _
        "total=" & CStr(totalItems) & vbLf & _
        "stage=" & stageName & vbLf

    statusText = VTDocumentImportStatusPrefix(stageName)
    If stageName = "inserting" Then
        statusText = statusText & _
            CStr(currentItem) & "/" & CStr(totalItems)
    ElseIf stageName = "complete" Then
        statusText = statusText & _
            CStr(totalItems) & "/" & CStr(totalItems)
    End If
    Application.StatusBar = statusText
    DoEvents
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub VTCancelWordDocumentImportDispatch( _
    ByVal dispatch As Object)

    Dim sourceDocumentId As String
    Dim bookmarkName As String

    VTRequireWritableWordDocument
    VTRequireDispatchValue dispatch, "sourceDocumentId"
    VTRequireDispatchValue dispatch, "bookmarkName"
    sourceDocumentId = CStr(dispatch("sourceDocumentId"))
    bookmarkName = CStr(dispatch("bookmarkName"))
    If sourceDocumentId <> VTWordDocumentIdentity() Then
        Err.Raise vbObjectError + 7585, "VisualTeX", _
            "The active Word document changed while the document importer was open."
    End If
    If Left$(bookmarkName, _
       Len(VT_WORD_DOCUMENT_IMPORT_BOOKMARK_PREFIX)) <> _
       VT_WORD_DOCUMENT_IMPORT_BOOKMARK_PREFIX Then
        Err.Raise vbObjectError + 7585, "VisualTeX", _
            "The document import Bookmark is invalid."
    End If
    If ActiveDocument.Bookmarks.Exists(bookmarkName) Then
        ActiveDocument.Bookmarks(bookmarkName).Delete
    End If
End Sub

Public Function VTDocumentImportSourceOffset( _
    ByVal manifest As Object, _
    ByVal itemIndex As Long, _
    ByVal suffix As String) As Long

    Dim rawValue As String
    Dim numericValue As Double

    rawValue = VTDocumentImportRequired(manifest, itemIndex, suffix)
    If Not IsNumeric(rawValue) Then GoTo InvalidOffset
    numericValue = CDbl(rawValue)
    If numericValue < 0# Or _
       numericValue > CDbl(VT_WORD_LATEX_REDRAW_MAX_BYTES) Or _
       numericValue <> Fix(numericValue) Then GoTo InvalidOffset
    VTDocumentImportSourceOffset = CLng(numericValue)
    Exit Function

InvalidOffset:
    Err.Raise vbObjectError + 7591, "VisualTeX", _
        "A LaTeX redraw source range is invalid."
End Function

Public Sub VTCommitWordDocumentImportDispatch( _
    ByVal sessionId As String, _
    ByVal dispatch As Object)

    Dim manifest As Object
    Dim insertedFormulaIds As Collection
    Dim targetDocument As Document
    Dim anchorBookmark As Bookmark
    Dim cursorRange As Range
    Dim rollbackRange As Range
    Dim sharedNativeDocument As Document
    Dim sharedVectorDocument As Document
    Dim vectorBatchDocumentPath As String
    Dim vectorBatchFormulaCount As Long
    Dim formulaId As Variant
    Dim manifestPath As String
    Dim nativeBatchDocumentPath As String
    Dim sourceDocumentId As String
    Dim bookmarkName As String
    Dim operationName As String
    Dim outputKind As String
    Dim itemKind As String
    Dim itemParagraphId As String
    Dim itemParagraphStyle As String
    Dim itemParagraphAlignment As String
    Dim itemListKind As String
    Dim activeParagraphId As String
    Dim activeParagraphStyle As String
    Dim activeParagraphAlignment As String
    Dim activeListKind As String
    Dim activeFormulaId As String
    Dim formulaFailureDescription As String
    Dim itemCount As Long
    Dim itemIndex As Long
    Dim nativeBatchFormulaCount As Long
    Dim itemListLevel As Long
    Dim activeListLevel As Long
    Dim insertionStart As Long
    Dim insertedEnd As Long
    Dim itemInsertionStart As Long
    Dim progressInterval As Long
    Dim activeParagraphStart As Long
    Dim itemParagraphStart As Boolean
    Dim itemParagraphEnd As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim transactionStage As String
    Dim internalMutationStarted As Boolean
    Dim previousScreenUpdating As Boolean
    Dim screenUpdatingChanged As Boolean
    Dim addedParagraphMark As Long

    addedParagraphMark = -1
    On Error GoTo Failed
    transactionStage = "commit-enter"
    VTTraceDocumentImportStage sessionId, "commit-enter", -1
    transactionStage = "commit-require-active-document"
    VTRequireWritableWordDocument
    transactionStage = "commit-validate-dispatch"
    VTRequireDispatchValue dispatch, "sourceDocumentId"
    VTRequireDispatchValue dispatch, "bookmarkName"
    VTRequireDispatchValue dispatch, "documentImportPath"
    transactionStage = "commit-read-dispatch"
    sourceDocumentId = CStr(dispatch("sourceDocumentId"))
    bookmarkName = CStr(dispatch("bookmarkName"))
    manifestPath = CStr(dispatch("documentImportPath"))
    transactionStage = "commit-resolve-document"
    Set targetDocument = VTFindOpenWordDocumentByIdentity(sourceDocumentId)
    If targetDocument Is Nothing Then
        Err.Raise vbObjectError + 7586, "VisualTeX", _
            "The Word document used to start this import is no longer open."
    End If
    transactionStage = "commit-activate-document"
    targetDocument.Activate
    transactionStage = "commit-require-target-document"
    VTRequireWritableWordDocument

    transactionStage = "commit-read-manifest"
    Set manifest = VTReadDocumentImportManifest( _
        manifestPath, sessionId)
    transactionStage = "commit-validate-manifest-target"
    If CStr(manifest("sourceDocumentId")) <> sourceDocumentId Or _
       CStr(manifest("bookmarkName")) <> bookmarkName Then
        Err.Raise vbObjectError + 7586, "VisualTeX", _
            "The document import manifest target does not match Word."
    End If
    operationName = CStr(manifest("operation"))

    transactionStage = "commit-check-bookmark"
    If Not targetDocument.Bookmarks.Exists(bookmarkName) Then
        If operationName = "formulaRestore" And _
           VTDispatchOptional(manifest, "redrawScope") = "document" Then
            transactionStage = "commit-recover-document-bookmark"
            Set rollbackRange = targetDocument.Content.Duplicate
            If StrComp( _
               rollbackRange.Text, _
               VTBase64UrlDecodeUtf8(VTDispatchOptional( _
                   manifest, "targetTextBase64")), _
               vbBinaryCompare) <> 0 Then
                Err.Raise vbObjectError + 7586, "VisualTeX", _
                    "The Word document changed while global formula conversion was running."
            End If
            targetDocument.Bookmarks.Add _
                Name:=bookmarkName, Range:=rollbackRange
        Else
            Err.Raise vbObjectError + 7586, "VisualTeX", _
                "The document import insertion point no longer exists."
        End If
    End If
    If Not targetDocument.Bookmarks.Exists(bookmarkName) Then
        Err.Raise vbObjectError + 7586, "VisualTeX", _
            "Word could not recover the document conversion target."
    End If
    If operationName <> "documentImport" And _
       operationName <> "latexRedraw" And _
       operationName <> "formulaRestore" Then
        Err.Raise vbObjectError + 7586, "VisualTeX", _
            "The document import manifest operation is invalid."
    End If
    outputKind = CStr(manifest("outputKind"))
    If operationName = "formulaRestore" Then
        If outputKind <> "latex" And outputKind <> "image" And _
           outputKind <> "omml" Then
            Err.Raise vbObjectError + 7586, "VisualTeX", _
                "The formula restore output kind is invalid."
        End If
    ElseIf outputKind <> "omml" And outputKind <> "image" Then
        Err.Raise vbObjectError + 7586, "VisualTeX", _
            "The document formula output kind is invalid."
    End If
    transactionStage = "commit-read-item-count"
    If Not IsNumeric(manifest("itemCount")) Then
        Err.Raise vbObjectError + 7586, "VisualTeX", _
            "The document import item count is invalid."
    End If
    itemCount = CLng(manifest("itemCount"))
    If operationName = "latexRedraw" Or _
       operationName = "formulaRestore" Then
        If itemCount < 1 Or _
           itemCount > VT_WORD_LATEX_REDRAW_MAX_FORMULAS Then
            Err.Raise vbObjectError + 7586, "VisualTeX", _
                "The Word range replacement count is outside the supported range."
        End If
        transactionStage = "commit-range-replacement"
        VTCommitWordLatexRedrawDispatch _
            sessionId, manifest, sourceDocumentId, bookmarkName, _
            outputKind, itemCount
        Exit Sub
    End If
    If itemCount < 1 Or itemCount > 2048 Then
        Err.Raise vbObjectError + 7586, "VisualTeX", _
            "The document import item count is outside the supported range."
    End If
    VTUpdateDocumentImportProgress sessionId, 0, itemCount, "inserting"
    progressInterval = itemCount \ 10
    If progressInterval < 1 Then progressInterval = 1

    Set anchorBookmark = targetDocument.Bookmarks(bookmarkName)
    insertionStart = anchorBookmark.Range.Start
    insertedEnd = insertionStart
    Set cursorRange = targetDocument.Range( _
        Start:=insertionStart, End:=insertionStart)
    anchorBookmark.Delete
    Set insertedFormulaIds = New Collection
    ' Word for Mac lays out and repaints the visible document after every
    ' paragraph, text run and formula the import writes. The LaTeX redraw
    ' already writes the same formulas with screen updating off.
    transactionStage = "import-disable-screen-updating"
    On Error Resume Next
    Err.Clear
    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    screenUpdatingChanged = (Err.Number = 0)
    Err.Clear
    On Error GoTo Failed
    VTBeginV2ImportPlanCache
    If outputKind = "omml" Then
        nativeBatchDocumentPath = VTDispatchOptional( _
            manifest, "nativeBatchDocumentPath")
        If Len(nativeBatchDocumentPath) > 0 Then
            If Not IsNumeric(VTDispatchOptional( _
               manifest, "nativeBatchFormulaCount")) Then
                Err.Raise vbObjectError + 7586, "VisualTeX", _
                    "The shared native-equation batch count is invalid."
            End If
            nativeBatchFormulaCount = CLng(VTDispatchOptional( _
                manifest, "nativeBatchFormulaCount"))
            If nativeBatchFormulaCount < 1 Or _
               nativeBatchFormulaCount > 1000 Then
                Err.Raise vbObjectError + 7586, "VisualTeX", _
                    "The shared native-equation batch count is outside the supported range."
            End If
            VTValidateAbsoluteVisualTeXPath nativeBatchDocumentPath
            If Not VTPathFileExists(nativeBatchDocumentPath) Then
                Err.Raise vbObjectError + 7586, "VisualTeX", _
                    "The shared native-equation staging document is missing."
            End If
            transactionStage = "open-shared-native-document"
            Set sharedNativeDocument = Documents.Open( _
                FileName:=nativeBatchDocumentPath, _
                ConfirmConversions:=False, _
                ReadOnly:=True, _
                AddToRecentFiles:=False, _
                Visible:=False)
            If sharedNativeDocument.OMaths.Count <> nativeBatchFormulaCount Then
                Err.Raise vbObjectError + 7586, "VisualTeX", _
                    "The shared native-equation staging document is incomplete."
            End If
            targetDocument.Activate
        End If
    End If
    If outputKind = "image" Then
        vectorBatchDocumentPath = VTDispatchOptional( _
            manifest, "vectorBatchDocumentPath")
        If Len(vectorBatchDocumentPath) > 0 Then
            vectorBatchFormulaCount = CLng(Val(VTDispatchOptional( _
                manifest, "vectorBatchFormulaCount")))
            If vectorBatchFormulaCount < 1 Or vectorBatchFormulaCount > 1000 Then
                Err.Raise vbObjectError + 7586, "VisualTeX", _
                    "The shared image-formula batch count is invalid."
            End If
            VTValidateAbsoluteVisualTeXPath vectorBatchDocumentPath
            If Not VTPathFileExists(vectorBatchDocumentPath) Then
                Err.Raise vbObjectError + 7586, "VisualTeX", _
                    "The shared image-formula staging document is missing."
            End If
            transactionStage = "open-shared-vector-document"
            Set sharedVectorDocument = Documents.Open( _
                FileName:=vectorBatchDocumentPath, _
                ConfirmConversions:=False, ReadOnly:=True, _
                AddToRecentFiles:=False, Visible:=False)
            If sharedVectorDocument.InlineShapes.Count <> vectorBatchFormulaCount Then
                Err.Raise vbObjectError + 7586, "VisualTeX", _
                    "The shared image-formula staging document is incomplete."
            End If
            targetDocument.Activate
        End If
    End If
    VTBeginWordInternalMutation
    internalMutationStarted = True

    ' Imported paragraphs do not join the paragraph the caret ends, as on Windows: at the
    ' end of a paragraph with text, the first one starts a new paragraph.
    If VTDocumentImportParagraphFlag(manifest, 0, "paragraphStart") Then
        transactionStage = "import-start-paragraph"
        Set rollbackRange = cursorRange.Paragraphs(1).Range.Duplicate
        If cursorRange.Start >= rollbackRange.End - 1 And _
           VTContainsVisibleWordBodyText(rollbackRange.Text) Then
            addedParagraphMark = cursorRange.Start
            cursorRange.InsertAfter vbCr
            cursorRange.Collapse wdCollapseEnd
        End If
        Set rollbackRange = Nothing
    End If

    For itemIndex = 0 To itemCount - 1
        VTTraceDocumentImportStage sessionId, "item-start", itemIndex
        itemInsertionStart = cursorRange.Start
        itemKind = VTDocumentImportRequired( _
            manifest, itemIndex, "kind")
        itemParagraphId = VTDocumentImportParagraphValue( _
            manifest, itemIndex, "paragraphId")
        itemParagraphStyle = VTDocumentImportParagraphValue( _
            manifest, itemIndex, "paragraphStyle")
        itemParagraphAlignment = VTDocumentImportParagraphValue( _
            manifest, itemIndex, "paragraphAlignment")
        itemListKind = VTDocumentImportParagraphValue( _
            manifest, itemIndex, "listKind")
        itemParagraphStart = VTDocumentImportParagraphFlag( _
            manifest, itemIndex, "paragraphStart")
        itemParagraphEnd = VTDocumentImportParagraphFlag( _
            manifest, itemIndex, "paragraphEnd")
        itemListLevel = 0
        If Len(VTDocumentImportParagraphValue( _
           manifest, itemIndex, "listLevel")) > 0 Then
            If Not IsNumeric(VTDocumentImportParagraphValue( _
               manifest, itemIndex, "listLevel")) Then
                Err.Raise vbObjectError + 7587, "VisualTeX", _
                    "The imported Word list level is invalid."
            End If
            itemListLevel = CLng(VTDocumentImportParagraphValue( _
                manifest, itemIndex, "listLevel"))
        End If

        If Len(itemParagraphId) > 0 Then
            If Not VTIsCanonicalUuid(itemParagraphId) Then
                Err.Raise vbObjectError + 7587, "VisualTeX", _
                    "The imported Word paragraph id is invalid."
            End If
            If itemParagraphStart Then
                If Len(activeParagraphId) > 0 Then
                    Err.Raise vbObjectError + 7587, "VisualTeX", _
                        "Imported Word paragraphs overlap."
                End If
                activeParagraphId = itemParagraphId
                activeParagraphStyle = itemParagraphStyle
                activeParagraphAlignment = itemParagraphAlignment
                activeListKind = itemListKind
                activeListLevel = itemListLevel
                activeParagraphStart = cursorRange.Start
                transactionStage = "paragraph-prepare item " & CStr(itemIndex)
                VTPrepareDocumentImportParagraph _
                    cursorRange, activeParagraphStyle, _
                    activeParagraphAlignment, activeListKind, activeListLevel
            ElseIf activeParagraphId <> itemParagraphId Then
                Err.Raise vbObjectError + 7587, "VisualTeX", _
                    "An imported Word paragraph continuation is missing its start."
            ElseIf activeParagraphStyle <> itemParagraphStyle Or _
                   activeParagraphAlignment <> itemParagraphAlignment Or _
                   activeListKind <> itemListKind Or _
                   activeListLevel <> itemListLevel Then
                Err.Raise vbObjectError + 7587, "VisualTeX", _
                    "An imported Word paragraph changed formatting mid-paragraph."
            End If
        ElseIf Len(activeParagraphId) > 0 Then
            Err.Raise vbObjectError + 7587, "VisualTeX", _
                "Imported paragraph content is missing paragraph metadata."
        ElseIf itemParagraphStart Or itemParagraphEnd Or _
               Len(itemParagraphStyle) > 0 Or _
               Len(itemParagraphAlignment) > 0 Or _
               Len(itemListKind) > 0 Or itemListLevel <> 0 Then
            Err.Raise vbObjectError + 7587, "VisualTeX", _
                "Imported paragraph metadata is incomplete."
        End If

        VTTraceDocumentImportStage sessionId, "item-metadata-resolved", itemIndex
        transactionStage = "item-content item " & CStr(itemIndex)
        Select Case itemKind
            Case "text"
                VTDocumentImportInsertText _
                    cursorRange, _
                    VTDocumentImportRequired( _
                        manifest, itemIndex, "textBase64")
            Case "formula"
                activeFormulaId = VTDocumentImportRequired( _
                    manifest, itemIndex, "formulaId")
                formulaFailureDescription = ""
                If Not sharedVectorDocument Is Nothing Then
                    If Not VTTryDocumentImportInsertFormula( _
                       cursorRange, manifest, itemIndex, outputKind, _
                       insertedFormulaIds, formulaFailureDescription, _
                       0#, False, sharedVectorDocument) Then
                        Err.Raise vbObjectError + 7584, "VisualTeX", _
                            "Formula import failed: " & formulaFailureDescription
                    End If
                ElseIf sharedNativeDocument Is Nothing Then
                    If Not VTTryDocumentImportInsertFormula( _
                       cursorRange, manifest, itemIndex, outputKind, _
                       insertedFormulaIds, formulaFailureDescription) Then
                        Err.Raise vbObjectError + 7584, "VisualTeX", _
                            "Formula import failed: " & formulaFailureDescription
                    End If
                Else
                    If Not VTTryDocumentImportInsertFormula( _
                       cursorRange, manifest, itemIndex, outputKind, _
                       insertedFormulaIds, formulaFailureDescription, _
                       0#, False, Empty, sharedNativeDocument) Then
                        Err.Raise vbObjectError + 7584, "VisualTeX", _
                            "Formula import failed: " & formulaFailureDescription
                    End If
                End If
            Case Else
                Err.Raise vbObjectError + 7586, "VisualTeX", _
                    "The document import contains an unsupported item."
        End Select
        VTTraceDocumentImportStage sessionId, "item-content-inserted", itemIndex
        If itemParagraphEnd Then
            If Len(activeParagraphId) = 0 Or _
               activeParagraphId <> itemParagraphId Then
                Err.Raise vbObjectError + 7587, "VisualTeX", _
                    "The imported Word paragraph end is invalid."
            End If
            VTTraceDocumentImportStage sessionId, "paragraph-finalize-start", itemIndex
            transactionStage = "paragraph-finalize item " & CStr(itemIndex)
            VTFinalizeDocumentImportParagraph _
                cursorRange, activeParagraphStart, activeParagraphStyle, _
                activeParagraphAlignment, activeListKind, activeListLevel, _
                sessionId, itemIndex
            activeParagraphId = ""
            activeParagraphStyle = ""
            activeParagraphAlignment = ""
            activeListKind = ""
            activeListLevel = 0
            VTTraceDocumentImportStage sessionId, "paragraph-finalize-complete", itemIndex
        End If
        insertedEnd = cursorRange.Start
        activeFormulaId = ""
        VTTraceDocumentImportStage sessionId, "item-complete", itemIndex
        If itemIndex + 1 = itemCount Or _
           (itemIndex + 1) Mod progressInterval = 0 Then
            VTUpdateDocumentImportProgress _
                sessionId, itemIndex + 1, itemCount, "inserting"
        End If
    Next itemIndex

    If Len(activeParagraphId) > 0 Then
        Err.Raise vbObjectError + 7587, "VisualTeX", _
            "The imported Word paragraph ended before its boundary."
    End If
    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    If Not sharedVectorDocument Is Nothing Then
        sharedVectorDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set sharedVectorDocument = Nothing
        If Not targetDocument Is Nothing Then targetDocument.Activate
    End If
    If Not sharedNativeDocument Is Nothing Then
        sharedNativeDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set sharedNativeDocument = Nothing
        targetDocument.Activate
    End If
    VTEndV2ImportPlanCache
    If screenUpdatingChanged Then
        Application.ScreenUpdating = previousScreenUpdating
        screenUpdatingChanged = False
    End If
    cursorRange.Collapse wdCollapseStart
    cursorRange.Select
    VTUpdateDocumentImportProgress _
        sessionId, itemCount, itemCount, "complete"
    Application.StatusBar = ""
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    VTEndV2ImportPlanCache
    If screenUpdatingChanged Then
        Application.ScreenUpdating = previousScreenUpdating
        screenUpdatingChanged = False
    End If
    VTUpdateDocumentImportProgress sessionId, itemIndex, itemCount, "error"
    ' Range-replacement failures are traced by VTCommitWordLatexRedrawDispatch
    ' with the exact redraw stage. Do not overwrite that useful inner trace with
    ' the generic hand-off stage from this outer dispatcher.
    If transactionStage <> "commit-range-replacement" Then
        VTWriteWordFailureTrace _
            sessionId, transactionStage, errorNumber, errorDescription
    End If
    If Not sharedVectorDocument Is Nothing Then
        sharedVectorDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set sharedVectorDocument = Nothing
        If Not targetDocument Is Nothing Then targetDocument.Activate
    End If
    If Not sharedNativeDocument Is Nothing Then
        sharedNativeDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set sharedNativeDocument = Nothing
        If Not targetDocument Is Nothing Then targetDocument.Activate
    End If

    ' Never delete content from earlier completed items. The old rollback used
    ' insertionStart..insertedEnd, which removed every successful item while
    ' leaving the current failed item behind. Remove only the current item's
    ' local text tail; formula-specific structure is cleaned by exact identity.
    If Not targetDocument Is Nothing And Not cursorRange Is Nothing Then
        If itemInsertionStart >= 0 And _
           cursorRange.Start > itemInsertionStart Then
            Set rollbackRange = targetDocument.Range( _
                Start:=itemInsertionStart, End:=cursorRange.Start)
            rollbackRange.Delete
        End If
    End If
    ' The paragraph started for the import goes again when its first item failed.
    If Not targetDocument Is Nothing And addedParagraphMark >= 0 And itemIndex = 0 Then
        Set rollbackRange = targetDocument.Range( _
            Start:=addedParagraphMark, End:=addedParagraphMark + 1)
        If rollbackRange.Text = vbCr Then rollbackRange.Delete
        itemInsertionStart = addedParagraphMark
    End If
    If Not targetDocument Is Nothing And _
       VTIsCanonicalUuid(activeFormulaId) Then
        ' The current item can fail before VTDocumentImportInsertFormula records
        ' it as successfully inserted. Clean only that exact in-flight formula;
        ' previously completed formula metadata must remain valid.
        VTDeleteEquationNumberScaffold _
            targetDocument, activeFormulaId, False
        VTDeleteDocumentImportedFormulaState _
            targetDocument, activeFormulaId
    End If
    If Not targetDocument Is Nothing Then
        If Not targetDocument.Bookmarks.Exists(bookmarkName) Then
            Set cursorRange = targetDocument.Range( _
                Start:=itemInsertionStart, End:=itemInsertionStart)
            targetDocument.Bookmarks.Add _
                Name:=bookmarkName, Range:=cursorRange
        End If
    End If
    If internalMutationStarted Then VTEndWordInternalMutation
    On Error GoTo 0
    Err.Raise errorNumber, "VisualTeX Word document import", _
        transactionStage & ": " & errorDescription
End Sub

Private Function VTPrepareBatchUnnumberedDisplayInsertionRange( _
    ByVal requestedRange As Range) As Range

    Dim insertionRange As Range
    Dim paragraphRange As Range

    If requestedRange Is Nothing Then
        Err.Raise vbObjectError + 7551, "VisualTeX", _
            "The batch display insertion Range is missing."
    End If
    Set insertionRange = requestedRange.Duplicate
    insertionRange.Collapse wdCollapseStart
    If VTRangeIsWithinTable(insertionRange) Then
        Set VTPrepareBatchUnnumberedDisplayInsertionRange = _
            VTPrepareWordCreateInsertionRange(insertionRange, "block")
        Exit Function
    End If
    Set paragraphRange = insertionRange.Paragraphs(1).Range.Duplicate
    If paragraphRange.Fields.Count <> 0 Or _
       paragraphRange.InlineShapes.Count <> 0 Or _
       paragraphRange.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(paragraphRange) Then
        Set VTPrepareBatchUnnumberedDisplayInsertionRange = _
            VTPrepareWordCreateInsertionRange(insertionRange, "block")
        Exit Function
    End If

    ' Document import already owns this empty paragraph. Do not normalize it as
    ' plain body text just before the formula insertion immediately rewrites the
    ' same paragraph as a display-equation paragraph.
    Set VTPrepareBatchUnnumberedDisplayInsertionRange = insertionRange
End Function

Private Function VTBatchContinuationRangeAfterUnnumberedDisplayFormula( _
    ByVal formulaRange As Range) As Range

    Dim documentObject As Document
    Dim exactRange As Range
    Dim sourceParagraph As Range
    Dim continuationParagraph As Range
    Dim probeRange As Range
    Dim insertionRange As Range
    Dim caretRange As Range
    Dim continuationStart As Long
    Dim markerStart As Long

    If formulaRange Is Nothing Or formulaRange.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "The batch display formula continuation target is invalid."
    End If
    Set documentObject = formulaRange.Document
    If formulaRange.OMaths.Count = 1 Then
        Set exactRange = VTNativeMathRoot(formulaRange.OMaths(1)).Range.Duplicate
    ElseIf formulaRange.InlineShapes.Count = 1 Then
        Set exactRange = formulaRange.InlineShapes(1).Range.Duplicate
    Else
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "The batch display formula continuation target is ambiguous."
    End If
    Set sourceParagraph = exactRange.Paragraphs(1).Range.Duplicate
    If exactRange.Start < sourceParagraph.Start Or _
       exactRange.End > sourceParagraph.End Then
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "The batch display formula left its dedicated paragraph."
    End If

    ' Batch import owns the paragraph immediately after an unnumbered display
    ' formula. Do not run the interactive continuation normalizer here: the next
    ' text item applies its paragraph style before inserting content, while a
    ' following formula prepares its own dedicated paragraph. Reuse an existing
    ' empty paragraph or create exactly one at the formula boundary.
    continuationStart = sourceParagraph.End
    If continuationStart < documentObject.Content.End Then
        Set probeRange = documentObject.Range( _
            Start:=continuationStart, End:=continuationStart + 1)
        Set continuationParagraph = probeRange.Paragraphs(1).Range.Duplicate
        If continuationParagraph.Start < continuationStart Or _
           continuationParagraph.Information(wdWithInTable) Or _
           continuationParagraph.Fields.Count <> 0 Or _
           continuationParagraph.InlineShapes.Count <> 0 Or _
           continuationParagraph.OMaths.Count <> 0 Or _
           VTWordRangeHasMeaningfulText(continuationParagraph) Then
            Set continuationParagraph = Nothing
        End If
    End If

    If continuationParagraph Is Nothing Then
        markerStart = sourceParagraph.End - 1
        If markerStart < sourceParagraph.Start Then
            Err.Raise vbObjectError + 7553, "VisualTeX", _
                "The batch display formula paragraph has no insertion boundary."
        End If
        Set insertionRange = documentObject.Range( _
            Start:=markerStart, End:=markerStart)
        insertionRange.Text = vbCr
        continuationStart = markerStart + 1
        Set probeRange = documentObject.Range( _
            Start:=continuationStart, End:=continuationStart + 1)
        Set continuationParagraph = probeRange.Paragraphs(1).Range.Duplicate
        If continuationParagraph.Start <> continuationStart Or _
           continuationParagraph.Information(wdWithInTable) Or _
           continuationParagraph.Fields.Count <> 0 Or _
           continuationParagraph.InlineShapes.Count <> 0 Or _
           continuationParagraph.OMaths.Count <> 0 Or _
           VTWordRangeHasMeaningfulText(continuationParagraph) Then
            Err.Raise vbObjectError + 7553, "VisualTeX", _
                "Word did not expose an empty batch continuation paragraph."
        End If
    End If

    Set caretRange = documentObject.Range( _
        Start:=continuationParagraph.Start, _
        End:=continuationParagraph.Start)
    Set VTBatchContinuationRangeAfterUnnumberedDisplayFormula = _
        caretRange.Duplicate
End Function

' --- numbering context of a running document import -------------------------

' A document import inserts in document order at its cursor and never changes
' what lies before the cursor, so the numbered equations and headings found
' there stay exactly what a new search of the document would find. Between
' VTBeginV2ImportPlanCache and VTEndV2ImportPlanCache, VTV2PlanInsertion reads
' only what was added since the import's previous plan: Word for Mac takes
' ~27 ms to read one numbered equation, so the k-th number of an import used
' to read all k - 1 earlier ones again.
Private Sub VTBeginV2ImportPlanCache()
    VTEndV2ImportPlanCache
    VT_WORD_V2_IMPORT_CACHE_ACTIVE = True
End Sub

Private Sub VTEndV2ImportPlanCache()
    VT_WORD_V2_IMPORT_CACHE_ACTIVE = False
    VT_WORD_V2_IMPORT_CACHE_DOCUMENT = ""
    VT_WORD_V2_IMPORT_CACHE_POSITION = -1
    Set VT_WORD_V2_IMPORT_ENTRIES = Nothing
    Set VT_WORD_V2_IMPORT_HEADINGS = Nothing
    VT_WORD_V2_IMPORT_HEADINGS_END = -1
    VT_WORD_V2_IMPORT_HEADINGS_LEVEL = -1
End Sub

' True when the cache was filled for doc at a position no later than position
' and its last equation and heading are still exactly where they were found.
' Anything that moved content before them makes the whole cache unusable.
Public Function VTV2ImportCacheUsable( _
    ByVal doc As Document, ByVal position As Long) As Boolean

    Dim entry As Variant
    Dim probe As Range

    If Not VT_WORD_V2_IMPORT_CACHE_ACTIVE Then Exit Function
    If Len(VT_WORD_V2_IMPORT_CACHE_DOCUMENT) = 0 Then Exit Function
    If position < VT_WORD_V2_IMPORT_CACHE_POSITION Then Exit Function
    On Error GoTo NotUsable
    If StrComp(doc.FullName, VT_WORD_V2_IMPORT_CACHE_DOCUMENT, vbBinaryCompare) <> 0 Then Exit Function
    If Not VT_WORD_V2_IMPORT_ENTRIES Is Nothing Then
        If VT_WORD_V2_IMPORT_ENTRIES.Count > 0 Then
            entry = VT_WORD_V2_IMPORT_ENTRIES(VT_WORD_V2_IMPORT_ENTRIES.Count)
            Set probe = doc.Range(entry(0), entry(8))
            If probe.OMaths.Count < 1 Then Exit Function
            Set probe = VTNativeMathRoot(probe.OMaths(1)).Range
            If probe.Start <> entry(0) Or probe.End <> entry(8) Then Exit Function
        End If
    End If
    If Not VT_WORD_V2_IMPORT_HEADINGS Is Nothing Then
        If VT_WORD_V2_IMPORT_HEADINGS.Count > 0 Then
            entry = VT_WORD_V2_IMPORT_HEADINGS(VT_WORD_V2_IMPORT_HEADINGS.Count)
            Set probe = doc.Range(entry(0), entry(0)).Paragraphs(1).Range
            If probe.Start <> entry(0) Then Exit Function
            If probe.ParagraphFormat.OutlineLevel <> entry(1) Then Exit Function
        End If
    End If
    VTV2ImportCacheUsable = True
    Exit Function

NotUsable:
    Err.Clear
End Function

Public Sub VTV2RememberImportPosition(ByVal doc As Document, ByVal position As Long)
    VT_WORD_V2_IMPORT_CACHE_DOCUMENT = doc.FullName
    VT_WORD_V2_IMPORT_CACHE_POSITION = position
End Sub
