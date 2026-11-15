Attribute VB_Name = "VTWordRedraw"
Option Explicit
Option Private Module

' LaTeX redraw (LaTeX text to formulas) and formula restore (formulas back to LaTeX text).

Private Const VT_WORD_LATEX_REDRAW_SOURCE_FILE As String = _
    "/latex-redraw-source.txt"
Private Const VT_WORD_FORMULA_RESTORE_SOURCE_FILE As String = _
    "/formula-restore-source.txt"
Private Const VT_WORD_LATEX_REDRAW_CHUNK_SIZE As Long = 16000
Private Const VT_WORD_FORMULA_RESTORE_COMBINED_MAX_BYTES As Long = 2097152

Private Function VTLatexRedrawRequestJson( _
    ByVal sessionId As String, _
    ByVal sourceDocumentId As String, _
    ByVal bookmarkName As String, _
    ByVal defaultFontSizePt As Double, _
    ByVal redrawScope As String, _
    ByVal outputKind As String) As String

    VTLatexRedrawRequestJson = VTDocumentImportRequestJson( _
        sessionId, sourceDocumentId, bookmarkName, defaultFontSizePt)
    VTLatexRedrawRequestJson = Replace$( _
        VTLatexRedrawRequestJson, _
        """operation"":""documentImport""", _
        """operation"":""latexRedraw""")
    VTLatexRedrawRequestJson = _
        Left$(VTLatexRedrawRequestJson, _
            Len(VTLatexRedrawRequestJson) - 2) & _
        "," & """redrawScope"":" & VTJsonString(redrawScope) & _
        "," & """outputKind"":" & VTJsonString(outputKind) & _
        "}}"
End Function

Private Function VTFormulaRestoreRequestJson( _
    ByVal sessionId As String, _
    ByVal sourceDocumentId As String, _
    ByVal bookmarkName As String, _
    ByVal defaultFontSizePt As Double, _
    ByVal restoreScope As String, _
    ByVal sourceKind As String, _
    ByVal outputKind As String) As String

    VTFormulaRestoreRequestJson = VTDocumentImportRequestJson( _
        sessionId, sourceDocumentId, bookmarkName, defaultFontSizePt)
    VTFormulaRestoreRequestJson = Replace$( _
        VTFormulaRestoreRequestJson, _
        """operation"":""documentImport""", _
        """operation"":""formulaRestore""")
    VTFormulaRestoreRequestJson = _
        Left$(VTFormulaRestoreRequestJson, _
            Len(VTFormulaRestoreRequestJson) - 2) & _
        "," & """redrawScope"":" & VTJsonString(restoreScope) & _
        "," & """outputKind"":" & VTJsonString(outputKind) & _
        "," & """sourceKind"":" & VTJsonString(sourceKind) & _
        "}}"
End Function

Private Sub VTWriteLargeSessionText( _
    ByVal targetPath As String, _
    ByVal contents As String)

    Dim sourceOffset As Long
    Dim sourceChunk As String

    sourceOffset = 1
    sourceChunk = Mid$( _
        contents, sourceOffset, VT_WORD_LATEX_REDRAW_CHUNK_SIZE)
    VTWriteTextAtomic targetPath, sourceChunk
    sourceOffset = sourceOffset + Len(sourceChunk)
    Do While sourceOffset <= Len(contents)
        sourceChunk = Mid$( _
            contents, sourceOffset, VT_WORD_LATEX_REDRAW_CHUNK_SIZE)
        VTAppendText targetPath, sourceChunk
        sourceOffset = sourceOffset + Len(sourceChunk)
    Loop
End Sub

Public Function VTFormulaRestoreSelectionRange( _
    ByVal sourceKind As String) As Range

    Dim probeRange As Range
    Dim formulaShape As InlineShape
    Dim formulaMath As OMath

    If sourceKind = "image" Then
        Set formulaShape = VTVisualTeXInlineShapeAtSelection(Selection)
        If formulaShape Is Nothing Then
            Err.Raise vbObjectError + 7595, "VisualTeX", _
                "Select one VisualTeX image formula."
        End If
        Set VTFormulaRestoreSelectionRange = _
            formulaShape.Range.Duplicate
        Exit Function
    End If

    Set probeRange = Selection.Range.Duplicate
    If probeRange.OMaths.Count = 0 Then
        On Error Resume Next
        probeRange.MoveStart Unit:=wdCharacter, Count:=-1
        probeRange.MoveEnd Unit:=wdCharacter, Count:=1
        Err.Clear
        On Error GoTo 0
    End If
    If probeRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7595, "VisualTeX", _
            "Select one Word native equation."
    End If
    Set formulaMath = probeRange.OMaths(1)
    If formulaMath.Range.InlineShapes.Count <> 0 Then
        Err.Raise vbObjectError + 7595, "VisualTeX", _
            "The selected math container contains an image, not a native Word equation."
    End If
    Set VTFormulaRestoreSelectionRange = formulaMath.Range.Duplicate
End Function

Private Function VTFormulaRestoreNumberingScaffoldOwnsRange( _
    ByVal sourceRange As Range, _
    ByVal formulaId As String) As Boolean

    Dim documentObject As Document
    Dim numberedRange As Range
    Dim exactSource As Range
    Dim exactNumbered As Range

    If sourceRange Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then
        Exit Function
    End If
    Set documentObject = sourceRange.Document
    If Not documentObject.Bookmarks.Exists( _
       VTEquationNumberBookmarkName(formulaId)) And _
       Not documentObject.Bookmarks.Exists( _
       VTEquationSequenceNumberBookmarkName(formulaId)) And _
       Not documentObject.Bookmarks.Exists( _
       VTEquationCaptionBookmarkName(formulaId)) Then Exit Function

    Set numberedRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If numberedRange Is Nothing Then Exit Function

    If sourceRange.OMaths.Count = 1 And _
       sourceRange.InlineShapes.Count = 0 Then
        Set exactSource = VTNativeMathRoot(sourceRange.OMaths(1)).Range.Duplicate
    ElseIf sourceRange.InlineShapes.Count = 1 And _
           sourceRange.OMaths.Count = 0 Then
        Set exactSource = sourceRange.InlineShapes(1).Range.Duplicate
    Else
        Exit Function
    End If

    If numberedRange.OMaths.Count = 1 And _
       numberedRange.InlineShapes.Count = 0 Then
        Set exactNumbered = VTNativeMathRoot(numberedRange.OMaths(1)).Range.Duplicate
    ElseIf numberedRange.InlineShapes.Count = 1 And _
           numberedRange.OMaths.Count = 0 Then
        Set exactNumbered = numberedRange.InlineShapes(1).Range.Duplicate
    Else
        Exit Function
    End If

    VTFormulaRestoreNumberingScaffoldOwnsRange = _
        (exactNumbered.Start = exactSource.Start And _
         exactNumbered.End = exactSource.End)
End Function

Private Function VTTryResolveFormulaRestoreNumbering( _
    ByVal sourceRange As Range, _
    ByRef formulaId As String, _
    ByRef legacyNumberTable As Table) As Boolean

    Dim formulaShape As InlineShape
    Dim nativeBookmark As Bookmark
    Dim displayMode As String
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean
    Dim numbered As Boolean
    Dim tableFormulaId As String

    formulaId = ""
    Set legacyNumberTable = Nothing
    If sourceRange Is Nothing Then Exit Function

    If sourceRange.InlineShapes.Count = 1 And _
       sourceRange.OMaths.Count = 0 Then
        Set formulaShape = sourceRange.InlineShapes(1)
        If Not VTTryResolveVisualTeXInlineShapeReference( _
           formulaShape, formulaId, displayMode, numbered, _
           encodedMetadata, metadataNeedsWrite, formatNeedsWrite) Then
            formulaId = ""
            Exit Function
        End If
    ElseIf sourceRange.InlineShapes.Count = 0 And _
           sourceRange.OMaths.Count = 1 Then
        Set nativeBookmark = VTFindNativeFormulaBookmark( _
            sourceRange.OMaths(1).Range, False)
        If nativeBookmark Is Nothing Then Exit Function
        If Not VTTryFormulaIdFromNativeBookmark( _
           nativeBookmark.Name, formulaId) Then
            formulaId = ""
            Exit Function
        End If
        If VTFormulaRestoreNumberingScaffoldOwnsRange( _
           sourceRange, formulaId) Then
            displayMode = "block"
            numbered = True
        ElseIf Not VTTryReadWordFormulaFormat( _
           sourceRange.Document, formulaId, displayMode, numbered) Then
            formulaId = ""
            Exit Function
        End If
    Else
        Exit Function
    End If

    If displayMode <> "block" Or Not numbered Or _
       Not VTIsCanonicalUuid(formulaId) Then
        formulaId = ""
        Exit Function
    End If

    If sourceRange.Information(wdWithInTable) Then
        Set legacyNumberTable = sourceRange.Tables(1)
        tableFormulaId = VTFormulaIdForValidNumberedTable( _
            legacyNumberTable)
        If StrComp(tableFormulaId, formulaId, vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 7595, "VisualTeX", _
                "The numbered formula is inside an unrecognized Word table."
        End If
    End If

    VTTryResolveFormulaRestoreNumbering = True
End Function

Private Function VTReplaceFormulaRestoreRangeWithLatex( _
    ByVal sourceRange As Range, _
    ByVal replacementText As String, _
    ByRef mutationOccurred As Boolean, _
    ByVal numberedFormulaIds As Collection, _
    ByVal numberedRestoreRanges As Collection) As Range

    Dim targetDocument As Document
    Dim insertionRange As Range
    Dim restoredRange As Range
    Dim legacyNumberTable As Table
    Dim formulaId As String
    Dim insertionStart As Long
    Dim wasNumbered As Boolean

    If sourceRange Is Nothing Or Len(replacementText) = 0 Then
        Err.Raise vbObjectError + 7595, "VisualTeX", _
            "The formula restore replacement is invalid."
    End If
    If numberedFormulaIds Is Nothing Or numberedRestoreRanges Is Nothing Then
        Err.Raise vbObjectError + 7595, "VisualTeX", _
            "The formula restore numbering plan is missing."
    End If
    Set targetDocument = sourceRange.Document
    wasNumbered = VTTryResolveFormulaRestoreNumbering( _
        sourceRange, formulaId, legacyNumberTable)

    If Not legacyNumberTable Is Nothing Then
        insertionStart = legacyNumberTable.Range.Start
        mutationOccurred = True
        legacyNumberTable.Delete
    Else
        insertionStart = sourceRange.Start
        mutationOccurred = True
        sourceRange.Delete
    End If

    Set insertionRange = targetDocument.Range( _
        Start:=insertionStart, End:=insertionStart)
    insertionRange.Text = replacementText
    Set restoredRange = targetDocument.Range( _
        Start:=insertionStart, _
        End:=insertionStart + Len(replacementText))

    If wasNumbered Then
        numberedFormulaIds.Add formulaId
        numberedRestoreRanges.Add restoredRange.Duplicate
    End If

    Set VTReplaceFormulaRestoreRangeWithLatex = restoredRange
End Function

Private Sub VTFinalizeFormulaRestoreNumbering( _
    ByVal documentObject As Document, _
    ByVal numberedFormulaIds As Collection, _
    ByVal numberedRestoreRanges As Collection)

    Dim restoredRange As Range
    Dim paragraphRange As Range
    Dim prefixRange As Range
    Dim suffixRange As Range
    Dim formulaId As String
    Dim paragraphContentEnd As Long
    Dim itemIndex As Long

    If documentObject Is Nothing Or numberedFormulaIds Is Nothing Or _
       numberedRestoreRanges Is Nothing Then Exit Sub
    If numberedFormulaIds.Count <> numberedRestoreRanges.Count Then
        Err.Raise vbObjectError + 7595, "VisualTeX", _
            "The formula restore numbering plan is inconsistent."
    End If

    ' Defer scaffold cleanup until every source formula has already been
    ' replaced. VTDeleteEquationNumberScaffold also freezes body REF fields;
    ' doing that between range replacements can shift Word positions for an
    ' earlier formula in the same bulk restore.
    For itemIndex = 1 To numberedFormulaIds.Count
        formulaId = CStr(numberedFormulaIds(itemIndex))
        VTDeleteEquationNumberScaffold documentObject, formulaId, False
    Next itemIndex

    For itemIndex = 1 To numberedRestoreRanges.Count
        Set restoredRange = numberedRestoreRanges(itemIndex)
        If restoredRange Is Nothing Then GoTo NextRange
        Set paragraphRange = restoredRange.Paragraphs(1).Range.Duplicate
        Set prefixRange = documentObject.Range( _
            Start:=paragraphRange.Start, End:=restoredRange.Start)
        If Not VTWordRangeHasMeaningfulText(prefixRange) Then
            prefixRange.Text = ""
        End If

        Set paragraphRange = restoredRange.Paragraphs(1).Range.Duplicate
        paragraphContentEnd = paragraphRange.End
        If paragraphContentEnd > paragraphRange.Start Then
            If documentObject.Range( _
               Start:=paragraphContentEnd - 1, _
               End:=paragraphContentEnd).Text = vbCr Then
                paragraphContentEnd = paragraphContentEnd - 1
            End If
        End If
        If restoredRange.End < paragraphContentEnd Then
            Set suffixRange = documentObject.Range( _
                Start:=restoredRange.End, End:=paragraphContentEnd)
            If Not VTWordRangeHasMeaningfulText(suffixRange) Then
                suffixRange.Text = ""
            End If
        End If
        Set paragraphRange = restoredRange.Paragraphs(1).Range.Duplicate
        VTNormalizePlainWordParagraph paragraphRange
NextRange:
    Next itemIndex
End Sub

Private Function VTCollectFormulaRestoreTargets( _
    ByVal targetRange As Range, _
    ByVal sourceKind As String) As Collection

    Dim targets As Collection
    Dim formulaShape As InlineShape
    Dim formulaMath As OMath
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean

    Set targets = New Collection
    If sourceKind = "omml" Then
        For Each formulaMath In VTOMathsInDocumentOrder(targetRange)
            If formulaMath.Range.InlineShapes.Count = 0 And Not VTIsEquationReferenceMath(formulaMath) Then
                targets.Add formulaMath
            End If
        Next formulaMath
    ElseIf sourceKind = "image" Then
        For Each formulaShape In targetRange.InlineShapes
            formulaId = ""
            displayMode = ""
            numbered = False
            encodedMetadata = ""
            metadataNeedsWrite = False
            formatNeedsWrite = False
            If VTTryResolveVisualTeXInlineShapeReference( _
               formulaShape, formulaId, displayMode, numbered, _
               encodedMetadata, metadataNeedsWrite, formatNeedsWrite) Then
                targets.Add formulaShape
            End If
        Next formulaShape
    Else
        Err.Raise vbObjectError + 7595, "VisualTeX", _
            "The Word formula restore source kind is invalid."
    End If
    Set VTCollectFormulaRestoreTargets = targets
End Function

Public Function VTFormulaRestoreNativeXml( _
    ByVal sourceRange As Range) As String

    Dim sourceXml As String
    Dim documentStart As Long
    Dim documentEnd As Long

    sourceXml = sourceRange.WordOpenXML
    ' WordOpenXML is a Flat OPC package: even one equation includes styles,
    ' theme, font tables and other package parts. Sending that entire package
    ' through VBA Base64/AppleScriptTask once per formula dominates batch time.
    ' Keep the current Word document part, including namespace declarations and
    ' the complete OMath/REF structure. No cached LaTeX or cached equation is used.
    documentStart = InStr(1, sourceXml, "<w:document ", vbBinaryCompare)
    If documentStart = 0 Then
        documentStart = InStr(1, sourceXml, "<w:document>", vbBinaryCompare)
    End If
    If documentStart > 0 Then
        documentEnd = InStr(documentStart, sourceXml, "</w:document>", vbBinaryCompare)
        If documentEnd > documentStart Then
            VTFormulaRestoreNativeXml = Mid$( _
                sourceXml, documentStart, _
                documentEnd - documentStart + Len("</w:document>"))
            Exit Function
        End If
    End If
    ' Retain compatibility if Word returns standalone OMML instead of Flat OPC.
    VTFormulaRestoreNativeXml = sourceXml
End Function

Private Function VTFormulaRestoreManifest( _
    ByVal sessionId As String, _
    ByVal sourceDocumentId As String, _
    ByVal bookmarkName As String, _
    ByVal targetRange As Range, _
    ByVal sourceKind As String, _
    ByVal targets As Collection, _
    ByVal outputKind As String) As String

    Dim manifest As String
    Dim prefix As String
    Dim sourceRange As Range
    Dim formulaShape As InlineShape
    Dim formulaMath As OMath
    Dim nativeBookmark As Bookmark
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean
    Dim fontSizePt As Double
    Dim payload As String
    Dim itemIndex As Long

    manifest = _
        "protocolVersion=" & CStr(VT_PROTOCOL_VERSION) & vbLf & _
        "sessionId=" & sessionId & vbLf & _
        "sourceDocumentId=" & _
            VTBase64UrlEncodeUtf8(sourceDocumentId) & vbLf & _
        "bookmarkName=" & bookmarkName & vbLf & _
        "sourceKind=" & sourceKind & vbLf & _
        "targetTextBase64=" & _
            VTBase64UrlEncodeUtf8(targetRange.Text) & vbLf & _
        "itemCount=" & CStr(targets.Count) & vbLf

    For itemIndex = 1 To targets.Count
        prefix = "item" & CStr(itemIndex - 1)
        If sourceKind = "omml" Then
            Set formulaMath = targets(itemIndex)
            Set sourceRange = formulaMath.Range.Duplicate
            formulaId = ""
            displayMode = ""
            numbered = False
            encodedMetadata = ""
            Set nativeBookmark = Nothing
            If VTTryFindNativeFormulaBookmarkLocally( _
               sourceRange, nativeBookmark) Then
                If VTTryFormulaIdFromNativeBookmark( _
                   nativeBookmark.Name, formulaId) Then
                    If Not VTTryReadWordFormulaFormat( _
                       sourceRange.Document, formulaId, _
                       displayMode, numbered) Or _
                       Not VTTryReadWordMetadataPayload( _
                           sourceRange.Document, formulaId, _
                           encodedMetadata) Then
                        formulaId = ""
                        displayMode = ""
                        numbered = False
                        encodedMetadata = ""
                    End If
                End If
            End If
            If Not VTPureNativeNumberField(sourceRange) Is Nothing Then numbered = True
            If Len(displayMode) = 0 Then
                If formulaMath.Type = wdOMathDisplay Then
                    displayMode = "block"
                Else
                    displayMode = "inline"
                End If
            End If
            fontSizePt = sourceRange.Font.Size
            If Not VTValidWordFormulaFontSize(fontSizePt) Then
                fontSizePt = VTPreferredWordFormulaFontSize(sourceRange)
            End If
            ' Word expands WordOpenXML for a partial OMath Range back to the
            ' entire OMath container. Pass the exact complete XML here; the
            ' native converter removes only the verified VisualTeX eqArr/#/VT_N_
            ' numbering wrapper before OMML -> MathML conversion.
            payload = VTFormulaRestoreNativeXml(sourceRange)
        Else
            Set formulaShape = targets(itemIndex)
            Set sourceRange = VTVisualTeXImageContainerRange(formulaShape)
            If sourceRange Is Nothing Then
                Err.Raise vbObjectError + 7595, "VisualTeX", _
                    "A selected image formula lost its VisualTeX container."
            End If
            formulaId = ""
            displayMode = ""
            numbered = False
            encodedMetadata = ""
            metadataNeedsWrite = False
            formatNeedsWrite = False
            If Not VTTryResolveVisualTeXInlineShapeReference( _
               formulaShape, formulaId, displayMode, numbered, _
               encodedMetadata, metadataNeedsWrite, formatNeedsWrite) Then
                Err.Raise vbObjectError + 7595, "VisualTeX", _
                    "A selected image formula lost its VisualTeX metadata."
            End If
            ' The size the image was drawn at. A numbered image paragraph has the
            ' larger number style, which is not the formula's size.
            fontSizePt = VTWordImageFormulaFontSize(formulaShape)
            If Not VTValidWordFormulaFontSize(fontSizePt) Then
                fontSizePt = VTPreferredWordFormulaFontSize( _
                    formulaShape.Range.Duplicate)
            End If
            payload = encodedMetadata
        End If
        ' Format conversion gives the formula the global size for its kind.
        If outputKind <> "latex" Then
            fontSizePt = VTConversionFontSize(displayMode, fontSizePt)
        End If
        manifest = manifest & _
            prefix & "sourceStart=" & _
                CStr(sourceRange.Start - targetRange.Start) & vbLf & _
            prefix & "sourceEnd=" & _
                CStr(sourceRange.End - targetRange.Start) & vbLf & _
            prefix & "sourceTextBase64=" & _
                VTBase64UrlEncodeUtf8(sourceRange.Text) & vbLf & _
            prefix & "displayMode=" & displayMode & vbLf & _
            prefix & "fontSizePt=" & VTJsonNumber(fontSizePt) & vbLf & _
            prefix & "formulaId=" & formulaId & vbLf & _
            prefix & "numbered=" & IIf(numbered, "1", "0") & vbLf & _
            prefix & "metadata=" & encodedMetadata & vbLf & _
            prefix & "payloadBase64=" & _
                VTBase64UrlEncodeUtf8(payload) & vbLf
    Next itemIndex

    If VTUtf8ByteLength(manifest) > 16777216 Then
        Err.Raise vbObjectError + 7595, "VisualTeX", _
            "The Word formula restore transfer exceeds 16 MB. Select a smaller range."
    End If
    VTFormulaRestoreManifest = manifest
End Function

Private Sub VTCommitSynchronousOmmlLatexResult( _
    ByVal manifest As Object, _
    ByVal sourceDocumentId As String, _
    ByVal bookmarkName As String, _
    ByVal itemCount As Long)

    Dim targetDocument As Document
    Dim anchorBookmark As Bookmark
    Dim targetRange As Range
    Dim sourceRange As Range
    Dim rollbackRange As Range
    Dim undoRecord As Object
    Dim numberedFormulaIds As Collection
    Dim numberedRestoreRanges As Collection
    Dim sourceStarts() As Long
    Dim sourceEnds() As Long
    Dim sourceTexts() As String
    Dim replacementTexts() As String
    Dim targetStart As Long
    Dim targetEnd As Long
    Dim previousEnd As Long
    Dim itemIndex As Long
    Dim originalSource As String
    Dim itemKind As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim transactionStage As String
    Dim internalMutationStarted As Boolean
    Dim undoRecordStarted As Boolean
    Dim mutationOccurred As Boolean
    Dim rollbackVerified As Boolean

    On Error GoTo Failed
    transactionStage = "validate-target"
    If itemCount < 1 Or _
       itemCount > VT_WORD_LATEX_REDRAW_MAX_FORMULAS Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The synchronous OMML result count is invalid."
    End If
    If sourceDocumentId <> VTWordDocumentIdentity() Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The active Word document changed during OMML conversion."
    End If
    Set targetDocument = ActiveDocument
    If Not targetDocument.Bookmarks.Exists(bookmarkName) Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The OMML conversion target range no longer exists."
    End If
    Set anchorBookmark = targetDocument.Bookmarks(bookmarkName)
    Set targetRange = anchorBookmark.Range.Duplicate
    targetStart = targetRange.Start
    targetEnd = targetRange.End
    originalSource = targetRange.Text
    If targetEnd <= targetStart Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The OMML conversion target range is empty."
    End If

    ReDim sourceStarts(0 To itemCount - 1)
    ReDim sourceEnds(0 To itemCount - 1)
    ReDim sourceTexts(0 To itemCount - 1)
    ReDim replacementTexts(0 To itemCount - 1)
    previousEnd = 0
    For itemIndex = 0 To itemCount - 1
        transactionStage = "preflight item " & CStr(itemIndex)
        itemKind = VTDocumentImportRequired( _
            manifest, itemIndex, "kind")
        If itemKind <> "text" Then
            Err.Raise vbObjectError + 7591, "VisualTeX", _
                "The synchronous OMML result contains a non-text item."
        End If
        sourceStarts(itemIndex) = VTDocumentImportSourceOffset( _
            manifest, itemIndex, "sourceStart")
        sourceEnds(itemIndex) = VTDocumentImportSourceOffset( _
            manifest, itemIndex, "sourceEnd")
        If sourceStarts(itemIndex) < previousEnd Or _
           sourceEnds(itemIndex) <= sourceStarts(itemIndex) Or _
           sourceEnds(itemIndex) > targetEnd - targetStart Then
            Err.Raise vbObjectError + 7591, "VisualTeX", _
                "The synchronous OMML source ranges overlap or exceed the target."
        End If
        sourceTexts(itemIndex) = VTBase64UrlDecodeUtf8( _
            VTDocumentImportRequired( _
                manifest, itemIndex, "sourceTextBase64"))
        replacementTexts(itemIndex) = VTBase64UrlDecodeUtf8( _
            VTDocumentImportRequired( _
                manifest, itemIndex, "textBase64"))
        If Len(replacementTexts(itemIndex)) = 0 Then
            Err.Raise vbObjectError + 7591, "VisualTeX", _
                "A synchronous OMML LaTeX result is empty."
        End If
        Set sourceRange = targetDocument.Range( _
            Start:=targetStart + sourceStarts(itemIndex), _
            End:=targetStart + sourceEnds(itemIndex))
        If StrComp( _
           sourceRange.Text, sourceTexts(itemIndex), _
           vbBinaryCompare) <> 0 Then
            Err.Raise vbObjectError + 7591, "VisualTeX", _
                "Word content changed during synchronous OMML conversion."
        End If
        previousEnd = sourceEnds(itemIndex)
    Next itemIndex
    VTWordPerformanceMark "batch-word-preflighted"

    transactionStage = "start-undo"
    Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
    undoRecord.StartCustomRecord "VisualTeX Restore Word Formulas"
    undoRecordStarted = True
    VTBeginWordInternalMutation
    internalMutationStarted = True
    Set numberedFormulaIds = New Collection
    Set numberedRestoreRanges = New Collection
    anchorBookmark.Delete

    For itemIndex = itemCount - 1 To 0 Step -1
        transactionStage = "replace item " & CStr(itemIndex)
        Set sourceRange = targetDocument.Range( _
            Start:=targetStart + sourceStarts(itemIndex), _
            End:=targetStart + sourceEnds(itemIndex))
        Set sourceRange = VTReplaceFormulaRestoreRangeWithLatex( _
            sourceRange, replacementTexts(itemIndex), mutationOccurred, _
            numberedFormulaIds, numberedRestoreRanges)
    Next itemIndex
    VTFinalizeFormulaRestoreNumbering _
        targetDocument, numberedFormulaIds, numberedRestoreRanges

    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    If undoRecordStarted Then
        undoRecord.EndCustomRecord
        undoRecordStarted = False
    End If
    Set sourceRange = targetDocument.Range( _
        Start:=targetStart + sourceStarts(0), _
        End:=targetStart + sourceStarts(0) + _
            Len(replacementTexts(0)))
    sourceRange.Select
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    If undoRecordStarted Then
        undoRecord.EndCustomRecord
        undoRecordStarted = False
    End If
    rollbackVerified = Not mutationOccurred
    If mutationOccurred Then
        CallByName Application, "Undo", VbMethod
        Err.Clear
        Set rollbackRange = targetDocument.Range( _
            Start:=targetStart, End:=targetStart + Len(originalSource))
        rollbackVerified = _
            StrComp(rollbackRange.Text, originalSource, _
                vbBinaryCompare) = 0
    End If
    If rollbackVerified And Not targetDocument Is Nothing Then
        If Not targetDocument.Bookmarks.Exists(bookmarkName) Then
            Set rollbackRange = targetDocument.Range( _
                Start:=targetStart, _
                End:=targetStart + Len(originalSource))
            targetDocument.Bookmarks.Add _
                Name:=bookmarkName, Range:=rollbackRange
        End If
    End If
    On Error GoTo 0
    If Not rollbackVerified Then
        Err.Raise vbObjectError + 7592, _
            "VisualTeX synchronous OMML restore", _
            transactionStage & ": " & errorDescription & _
            " Word could not verify the automatic rollback."
    End If
    Err.Raise errorNumber, _
        "VisualTeX synchronous OMML restore", _
        transactionStage & ": " & errorDescription
End Sub

Public Function VTTryWriteOmmlBatchManifestDirect( _
    ByVal sessionId As String, _
    ByVal manifest As String) As Boolean

    Dim sessionPath As String
    Dim targetPath As String
    Dim temporaryPath As String
    Dim fileNumber As Integer
    Dim fileOpened As Boolean

    If Not VTIsCanonicalUuid(sessionId) Or Len(manifest) = 0 Then _
        Exit Function
    sessionPath = VTSessionDirectory(sessionId)
    targetPath = sessionPath & VT_WORD_FORMULA_RESTORE_SOURCE_FILE
    temporaryPath = targetPath & ".tmp"

    On Error GoTo DirectWriteUnavailable
    If Len(Dir$(sessionPath, vbDirectory)) = 0 Then MkDir sessionPath
    On Error Resume Next
    Kill temporaryPath
    Err.Clear
    On Error GoTo DirectWriteUnavailable
    fileNumber = FreeFile
    Open temporaryPath For Output Access Write As #fileNumber
    fileOpened = True
    Print #fileNumber, manifest;
    Close #fileNumber
    fileOpened = False
    On Error Resume Next
    Kill targetPath
    Err.Clear
    On Error GoTo DirectWriteUnavailable
    Name temporaryPath As targetPath
    VTTryWriteOmmlBatchManifestDirect = True
    Exit Function

DirectWriteUnavailable:
    On Error Resume Next
    If fileOpened Then Close #fileNumber
    Kill temporaryPath
    Err.Clear
    On Error GoTo 0
End Function

Private Function VTTryRestoreOmmlBatchSynchronously( _
    ByVal sessionId As String, _
    ByVal sourceDocumentId As String, _
    ByVal bookmarkName As String, _
    ByVal manifest As String) As Boolean

    Dim timingDetail As String
    Dim encodedResult As String
    Dim resultText As String
    Dim resultSeparator As Long
    Dim resultManifest As Object
    Dim itemCount As Long
    Dim directManifestWritten As Boolean
    Dim conversionErrorNumber As Long
    Dim commitErrorNumber As Long
    Dim commitErrorDescription As String

    If Len(manifest) > _
       VT_WORD_FORMULA_RESTORE_COMBINED_MAX_BYTES Then Exit Function

    VTWordPerformanceMark "batch-enter"
    directManifestWritten = _
        VTTryWriteOmmlBatchManifestDirect(sessionId, manifest)
    VTWordPerformanceMark "batch-source-ready"
    On Error GoTo ConversionUnavailable
    If directManifestWritten Then
        timingDetail = VTConvertOmmlBatch(VT_WORD_HOST, sessionId)
    Else
        timingDetail = VTConvertOmmlBatch( _
            VT_WORD_HOST, sessionId, manifest)
    End If
    VTWordPerformanceMark "batch-converted"
    resultSeparator = InStrRev( _
        timingDetail, "|", -1, vbBinaryCompare)
    If resultSeparator <= 1 Or _
       resultSeparator >= Len(timingDetail) Then
        Err.Raise vbObjectError + 7595, "VisualTeX", _
            "The synchronous OMML converter returned no result payload."
    End If
    encodedResult = Mid$(timingDetail, resultSeparator + 1)
    timingDetail = Left$(timingDetail, resultSeparator - 1)
    resultText = VTBase64UrlDecodeUtf8(encodedResult)
    Set resultManifest = VTParseDocumentImportManifestText( _
        resultText, sessionId)
    VTWordPerformanceMark "batch-result-decoded"
    If CStr(resultManifest("operation")) <> "formulaRestore" Or _
       CStr(resultManifest("outputKind")) <> "latex" Or _
       CStr(resultManifest("sourceDocumentId")) <> sourceDocumentId Or _
       CStr(resultManifest("bookmarkName")) <> bookmarkName Then
        Err.Raise vbObjectError + 7595, "VisualTeX", _
            "The synchronous OMML conversion result identity is invalid."
    End If
    itemCount = CLng(resultManifest("itemCount"))
    If itemCount < 1 Or _
       itemCount > VT_WORD_LATEX_REDRAW_MAX_FORMULAS Then
        Err.Raise vbObjectError + 7595, "VisualTeX", _
            "The synchronous OMML conversion item count is invalid."
    End If

    On Error GoTo CommitFailed
    VTCommitSynchronousOmmlLatexResult _
        resultManifest, sourceDocumentId, bookmarkName, itemCount
    VTWordPerformanceMark "batch-word-committed"
    If VT_WORD_PERFORMANCE_ACTIVE Then
        VTWordPerformanceFlush sessionId
    Else
        VTDeleteSessionFiles sessionId
    End If
    VTTryRestoreOmmlBatchSynchronously = True
    Exit Function

ConversionUnavailable:
    conversionErrorNumber = Err.Number
    Err.Clear
    If conversionErrorNumber = 0 Then conversionErrorNumber = 1
    VTTryRestoreOmmlBatchSynchronously = False
    Exit Function

CommitFailed:
    commitErrorNumber = Err.Number
    commitErrorDescription = Err.Description
    Err.Raise commitErrorNumber, _
        "VisualTeX synchronous OMML restore", _
        commitErrorDescription
End Function

Private Function VTFormulaRestoreNothingFoundMessage( _
    ByVal restoreScope As String, _
    ByVal sourceKind As String) As String

    Dim message As String

    ' The Windows add-in's English wording. Word's MsgBox on the Mac cannot
    ' display Chinese characters, so the Mac add-in's messages stay English.
    If restoreScope = "document" Then
        message = "The current Word document contains no convertible "
    Else
        message = "The current selection contains no convertible "
    End If
    If sourceKind = "omml" Then
        message = message & "Word OMML formulas."
    Else
        message = message & "VisualTeX image formulas."
    End If
    VTFormulaRestoreNothingFoundMessage = message
End Function

Public Sub VTStartWordFormulaRestore( _
    ByVal restoreScope As String, _
    ByVal sourceKind As String, _
    ByVal outputKind As String)

    Dim sessionId As String
    Dim bookmarkName As String
    Dim sourceDocumentId As String
    Dim targetRange As Range
    Dim targets As Collection
    Dim defaultFontSizePt As Double
    Dim requestJson As String
    Dim manifest As String
    Dim manifestPath As String
    Dim launchTiming As String
    Dim operationStage As String
    Dim errorNumber As Long
    Dim errorDescription As String

    If restoreScope <> "selection" And restoreScope <> "document" Then
        VTShowError "Word formula restore", vbObjectError + 7595, _
            "The formula restore scope is invalid."
        Exit Sub
    End If
    If sourceKind <> "omml" And sourceKind <> "image" Then
        VTShowError "Word formula restore", vbObjectError + 7595, _
            "The formula restore source kind is invalid."
        Exit Sub
    End If
    If outputKind <> "latex" And outputKind <> "image" And _
       outputKind <> "omml" Then
        VTShowError "Word formula restore", vbObjectError + 7595, _
            "The formula restore output kind is invalid."
        Exit Sub
    End If
    If outputKind = "image" And sourceKind <> "omml" Then
        VTShowError "Word formula restore", vbObjectError + 7595, _
            "Only Word OMML formulas can be converted to image formulas."
        Exit Sub
    End If
    If outputKind = "omml" And sourceKind <> "image" Then
        VTShowError "Word formula restore", vbObjectError + 7595, _
            "Only VisualTeX image formulas can be converted to Word OMML."
        Exit Sub
    End If

    VTWordPerformanceStartLocal sourceKind & "-to-" & outputKind
    On Error GoTo Failed
    operationStage = "require-document"
    VTRequireWritableWordDocument
    VTWordPerformanceMark "document-required"
    operationStage = "resolve-range"
    If restoreScope = "selection" And Selection.Range.Start = _
       Selection.Range.End Then
        Set targetRange = VTFormulaRestoreSelectionRange(sourceKind)
    ElseIf restoreScope = "selection" Then
        Set targetRange = Selection.Range.Duplicate
    Else
        Set targetRange = ActiveDocument.Content.Duplicate
    End If
    If targetRange Is Nothing Then
        Err.Raise vbObjectError + 7595, "VisualTeX", _
            "The Word formula restore range is missing."
    End If
    If targetRange.Start = targetRange.End Then
        Err.Raise vbObjectError + 7595, "VisualTeX", _
            "The Word formula restore range is empty."
    End If
    VTWordPerformanceMark "range-resolved"
    If VTUtf8ByteLength(targetRange.Text) > _
       VT_WORD_LATEX_REDRAW_MAX_BYTES Then
        Err.Raise vbObjectError + 7595, "VisualTeX", _
            "The Word formula restore range exceeds 5 MB."
    End If

    operationStage = "collect-formulas"
    Set targets = VTCollectFormulaRestoreTargets(targetRange, sourceKind)
    VTWordPerformanceMark "formulas-collected"
    If targets.Count = 0 Then
        ' Nothing to convert is information, not a failure (as on Windows).
        VTShowInformation VTFormulaRestoreNothingFoundMessage( _
            restoreScope, sourceKind)
        Exit Sub
    End If
    If targets.Count > VT_WORD_LATEX_REDRAW_MAX_FORMULAS Then
        Err.Raise vbObjectError + 7595, "VisualTeX", _
            "The target range contains more than 1000 formulas."
    End If

    operationStage = "create-identities"
    sessionId = VTNewUuidV4()
    bookmarkName = VTDocumentImportBookmarkName(sessionId)
    sourceDocumentId = VTWordDocumentIdentity()
    defaultFontSizePt = VTPreferredWordFormulaFontSize(targetRange)
    If ActiveDocument.Bookmarks.Exists(bookmarkName) Then
        ActiveDocument.Bookmarks(bookmarkName).Delete
    End If
    ActiveDocument.Bookmarks.Add Name:=bookmarkName, Range:=targetRange
    VTWordPerformanceMark "identities-created"

    operationStage = "serialize-request"
    requestJson = VTFormulaRestoreRequestJson( _
        sessionId, sourceDocumentId, bookmarkName, defaultFontSizePt, _
        restoreScope, sourceKind, outputKind)
    requestJson = VTV2WithDocumentMathFont(requestJson, ActiveDocument)
    VTWordPerformanceMark "request-serialized"
    manifest = VTFormulaRestoreManifest( _
        sessionId, sourceDocumentId, bookmarkName, targetRange, _
        sourceKind, targets, outputKind)
    VTWordPerformanceMark "manifest-serialized"
    If sourceKind = "omml" And outputKind = "latex" Then
        operationStage = "convert-omml-batch"
        If VTTryRestoreOmmlBatchSynchronously( _
           sessionId, sourceDocumentId, bookmarkName, manifest) Then
            operationStage = "complete"
            Exit Sub
        End If
    End If
    If VTUtf8ByteLength(manifest) <= _
       VT_WORD_FORMULA_RESTORE_COMBINED_MAX_BYTES Then
        operationStage = "write-source-and-launch"
        launchTiming = VTWriteFormulaRestoreAndLaunchSession( _
            VT_WORD_HOST, sessionId, requestJson, manifest)
    Else
        operationStage = "write-request"
        VTWriteRequest sessionId, requestJson
        manifestPath = VTSessionDirectory(sessionId) & _
            VT_WORD_FORMULA_RESTORE_SOURCE_FILE
        operationStage = "write-source"
        VTWriteLargeSessionText manifestPath, manifest
        operationStage = "launch"
        VTLaunchSession VT_WORD_HOST, sessionId
    End If
    VTWordPerformanceNote "launcherTiming", launchTiming
    VTWordPerformanceMark "restore-launch-complete"
    VTWordPerformanceFlush sessionId, 0, "", "/word-vba-scan-performance.txt"
    operationStage = "complete"
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Len(bookmarkName) > 0 Then
        If ActiveDocument.Bookmarks.Exists(bookmarkName) Then
            ActiveDocument.Bookmarks(bookmarkName).Delete
        End If
    End If
    If Len(sessionId) > 0 Then VTDeleteSessionFiles sessionId
    On Error GoTo 0
    VTShowError _
        "Word formula restore (" & operationStage & ")", _
        errorNumber, errorDescription
End Sub

Public Function VTRestoreCachedImageFormulasToLatex( _
    ByVal targetRange As Range) As Boolean

    Dim targetDocument As Document
    Dim formulaShape As InlineShape
    Dim containerRange As Range
    Dim replacementRange As Range
    Dim undoRecord As Object
    Dim numberedFormulaIds As Collection
    Dim numberedRestoreRanges As Collection
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim latexBase64 As String
    Dim latexText As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean
    Dim sourceStarts() As Long
    Dim sourceEnds() As Long
    Dim replacementTexts() As String
    Dim maximumCount As Long
    Dim formulaCount As Long
    Dim itemIndex As Long
    Dim internalMutationStarted As Boolean
    Dim undoRecordStarted As Boolean
    Dim mutationOccurred As Boolean
    Dim restoreErrorNumber As Long
    Dim restoreErrorDescription As String

    If targetRange Is Nothing Then Exit Function
    maximumCount = targetRange.InlineShapes.Count
    If maximumCount = 0 Then Exit Function
    Set targetDocument = targetRange.Document
    ReDim sourceStarts(1 To maximumCount)
    ReDim sourceEnds(1 To maximumCount)
    ReDim replacementTexts(1 To maximumCount)

    ' Preflight every target before changing Word. Old image formulas that do
    ' not have a durable LaTeX payload use the existing hidden restore pipeline
    ' instead of partially converting a document through this fast path.
    For Each formulaShape In targetRange.InlineShapes
        formulaId = ""
        displayMode = ""
        numbered = False
        encodedMetadata = ""
        metadataNeedsWrite = False
        formatNeedsWrite = False
        If VTTryResolveVisualTeXInlineShapeReference( _
           formulaShape, formulaId, displayMode, numbered, _
           encodedMetadata, metadataNeedsWrite, formatNeedsWrite) Then
            If Not VTTryReadWordLatexPayload( _
               targetDocument, formulaId, latexBase64) Then Exit Function
            latexText = VTBase64UrlDecodeUtf8(latexBase64)
            If Len(latexText) = 0 Then Exit Function
            Set containerRange = _
                VTVisualTeXImageContainerRange(formulaShape)
            If containerRange.Start >= targetRange.Start And _
               containerRange.End <= targetRange.End Then
                formulaCount = formulaCount + 1
                sourceStarts(formulaCount) = containerRange.Start
                sourceEnds(formulaCount) = containerRange.End
                If displayMode = "block" Then
                    replacementTexts(formulaCount) = _
                        "$$" & latexText & "$$"
                Else
                    replacementTexts(formulaCount) = _
                        "$" & latexText & "$"
                End If
            End If
        End If
    Next formulaShape
    If formulaCount = 0 Then Exit Function

    On Error GoTo RestoreFailed
    Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
    undoRecord.StartCustomRecord "VisualTeX Restore Image Formulas"
    undoRecordStarted = True
    VTBeginWordInternalMutation
    internalMutationStarted = True
    Set numberedFormulaIds = New Collection
    Set numberedRestoreRanges = New Collection
    For itemIndex = formulaCount To 1 Step -1
        Set replacementRange = targetDocument.Range( _
            Start:=sourceStarts(itemIndex), _
            End:=sourceEnds(itemIndex))
        Set replacementRange = VTReplaceFormulaRestoreRangeWithLatex( _
            replacementRange, replacementTexts(itemIndex), mutationOccurred, _
            numberedFormulaIds, numberedRestoreRanges)
    Next itemIndex
    VTFinalizeFormulaRestoreNumbering _
        targetDocument, numberedFormulaIds, numberedRestoreRanges
    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    If undoRecordStarted Then
        undoRecord.EndCustomRecord
        undoRecordStarted = False
    End If
    VTRestoreCachedImageFormulasToLatex = True
    Exit Function

RestoreFailed:
    restoreErrorNumber = Err.Number
    restoreErrorDescription = Err.Description
    On Error Resume Next
    If internalMutationStarted Then VTEndWordInternalMutation
    If undoRecordStarted Then undoRecord.EndCustomRecord
    If mutationOccurred Then CallByName Application, "Undo", VbMethod
    On Error GoTo 0
    Err.Raise restoreErrorNumber, "VisualTeX image-to-LaTeX", _
        restoreErrorDescription
End Function

Public Function VTRestoreCachedNativeFormulasToLatex( _
    ByVal targetRange As Range) As Boolean

    Dim targetDocument As Document
    Dim formulaMath As OMath
    Dim nativeBookmark As Bookmark
    Dim replacementRange As Range
    Dim undoRecord As Object
    Dim numberedFormulaIds As Collection
    Dim numberedRestoreRanges As Collection
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim latexBase64 As String
    Dim latexText As String
    Dim sourceStarts() As Long
    Dim sourceEnds() As Long
    Dim replacementTexts() As String
    Dim maximumCount As Long
    Dim formulaCount As Long
    Dim itemIndex As Long
    Dim internalMutationStarted As Boolean
    Dim undoRecordStarted As Boolean
    Dim mutationOccurred As Boolean
    Dim restoreErrorNumber As Long
    Dim restoreErrorDescription As String

    If targetRange Is Nothing Then Exit Function
    Set targetDocument = targetRange.Document
    If Not VTDocumentHasVisualTeXNativeBookmark(targetDocument) Then _
        Exit Function
    maximumCount = targetRange.OMaths.Count
    If maximumCount = 0 Then Exit Function
    ReDim sourceStarts(1 To maximumCount)
    ReDim sourceEnds(1 To maximumCount)
    ReDim replacementTexts(1 To maximumCount)

    ' Preflight every native formula. A Word-native or modified formula has no
    ' matching VisualTeX structure signature and therefore falls back to the
    ' reviewed OMML -> MathML -> LaTeX path without partial document changes.
    For Each formulaMath In VTOMathsInDocumentOrder(targetRange)
        If formulaMath.Range.InlineShapes.Count <> 0 Then Exit Function
        Set nativeBookmark = VTFindNativeFormulaBookmark( _
            formulaMath.Range, False)
        If nativeBookmark Is Nothing Then Exit Function
        If Not VTTryFormulaIdFromNativeBookmark( _
           nativeBookmark.Name, formulaId) Then Exit Function
        If Not VTTryReadWordFormulaFormat( _
           targetDocument, formulaId, displayMode, numbered) Then Exit Function
        If Not VTWordNativeSignatureMatches( _
           targetDocument, formulaId, formulaMath) Then Exit Function
        If Not VTTryReadWordLatexPayload( _
           targetDocument, formulaId, latexBase64) Then Exit Function
        latexText = VTBase64UrlDecodeUtf8(latexBase64)
        If Len(latexText) = 0 Then Exit Function
        formulaCount = formulaCount + 1
        sourceStarts(formulaCount) = formulaMath.Range.Start
        sourceEnds(formulaCount) = formulaMath.Range.End
        If displayMode = "block" Then
            replacementTexts(formulaCount) = "$$" & latexText & "$$"
        Else
            replacementTexts(formulaCount) = "$" & latexText & "$"
        End If
    Next formulaMath
    If formulaCount = 0 Then Exit Function

    On Error GoTo RestoreFailed
    Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
    undoRecord.StartCustomRecord "VisualTeX Restore OMML Formulas"
    undoRecordStarted = True
    VTBeginWordInternalMutation
    internalMutationStarted = True
    Set numberedFormulaIds = New Collection
    Set numberedRestoreRanges = New Collection
    For itemIndex = formulaCount To 1 Step -1
        Set replacementRange = targetDocument.Range( _
            Start:=sourceStarts(itemIndex), _
            End:=sourceEnds(itemIndex))
        Set replacementRange = VTReplaceFormulaRestoreRangeWithLatex( _
            replacementRange, replacementTexts(itemIndex), mutationOccurred, _
            numberedFormulaIds, numberedRestoreRanges)
    Next itemIndex
    VTFinalizeFormulaRestoreNumbering _
        targetDocument, numberedFormulaIds, numberedRestoreRanges
    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    If undoRecordStarted Then
        undoRecord.EndCustomRecord
        undoRecordStarted = False
    End If
    VTRestoreCachedNativeFormulasToLatex = True
    Exit Function

RestoreFailed:
    restoreErrorNumber = Err.Number
    restoreErrorDescription = Err.Description
    On Error Resume Next
    If internalMutationStarted Then VTEndWordInternalMutation
    If undoRecordStarted Then undoRecord.EndCustomRecord
    If mutationOccurred Then CallByName Application, "Undo", VbMethod
    On Error GoTo 0
    Err.Raise restoreErrorNumber, "VisualTeX OMML-to-LaTeX", _
        restoreErrorDescription
End Function

Private Sub VTWriteLatexRedrawSource( _
    ByVal sessionId As String, _
    ByVal sourceText As String)

    Dim sourcePath As String
    Dim sourceOffset As Long
    Dim sourceChunk As String

    If Len(sourceText) = 0 Or _
       VTUtf8ByteLength(sourceText) > VT_WORD_LATEX_REDRAW_MAX_BYTES Then
        Err.Raise vbObjectError + 7590, "VisualTeX", _
            "The LaTeX redraw source must be between 1 byte and 5 MB."
    End If

    sourcePath = VTSessionDirectory(sessionId) & _
        VT_WORD_LATEX_REDRAW_SOURCE_FILE
    sourceOffset = 1
    sourceChunk = Mid$( _
        sourceText, sourceOffset, VT_WORD_LATEX_REDRAW_CHUNK_SIZE)
    VTWriteTextAtomic sourcePath, sourceChunk
    sourceOffset = sourceOffset + Len(sourceChunk)
    Do While sourceOffset <= Len(sourceText)
        sourceChunk = Mid$( _
            sourceText, sourceOffset, VT_WORD_LATEX_REDRAW_CHUNK_SIZE)
        VTAppendText sourcePath, sourceChunk
        sourceOffset = sourceOffset + Len(sourceChunk)
    Loop
End Sub

Private Function VTUnsignedUtf16CodeUnit( _
    ByVal character As String) As Long

    Dim value As Long

    If Len(character) = 0 Then Exit Function
    value = AscW(character)
    If value < 0 Then value = value + 65536
    VTUnsignedUtf16CodeUnit = value
End Function

Private Function VTAdvanceWordOffsetToUtf16Boundary( _
    ByVal sourceText As String, _
    ByVal targetUtf16Offset As Long, _
    ByRef utf16Cursor As Long, _
    ByRef wordCursor As Long) As Long

    Dim firstUnit As Long
    Dim secondUnit As Long

    If targetUtf16Offset < utf16Cursor Or _
       targetUtf16Offset < 0 Or _
       targetUtf16Offset > Len(sourceText) Then
        Err.Raise vbObjectError + 7590, "VisualTeX", _
            "A LaTeX redraw UTF-16 offset is outside the source snapshot."
    End If

    Do While utf16Cursor < targetUtf16Offset
        firstUnit = VTUnsignedUtf16CodeUnit( _
            Mid$(sourceText, utf16Cursor + 1, 1))
        If firstUnit >= &HD800& And firstUnit <= &HDBFF& And _
           utf16Cursor + 1 < Len(sourceText) Then
            secondUnit = VTUnsignedUtf16CodeUnit( _
                Mid$(sourceText, utf16Cursor + 2, 1))
            If secondUnit >= &HDC00& And secondUnit <= &HDFFF& Then
                If utf16Cursor + 2 > targetUtf16Offset Then
                    Err.Raise vbObjectError + 7590, "VisualTeX", _
                        "A LaTeX redraw range splits a Unicode surrogate pair."
                End If
                utf16Cursor = utf16Cursor + 2
                wordCursor = wordCursor + 1
            Else
                utf16Cursor = utf16Cursor + 1
                wordCursor = wordCursor + 1
            End If
        Else
            utf16Cursor = utf16Cursor + 1
            wordCursor = wordCursor + 1
        End If
    Loop

    VTAdvanceWordOffsetToUtf16Boundary = wordCursor
End Function

Private Sub VTBuildWordOffsetsFromUtf16( _
    ByVal sourceText As String, _
    ByRef sourceStarts() As Long, _
    ByRef sourceEnds() As Long, _
    ByVal itemCount As Long, _
    ByRef wordStarts() As Long, _
    ByRef wordEnds() As Long)

    Dim utf16Cursor As Long
    Dim wordCursor As Long
    Dim itemIndex As Long

    If itemCount < 1 Then
        Err.Raise vbObjectError + 7590, "VisualTeX", _
            "LaTeX redraw requires at least one source range."
    End If

    ReDim wordStarts(0 To itemCount - 1)
    ReDim wordEnds(0 To itemCount - 1)
    utf16Cursor = 0
    wordCursor = 0

    For itemIndex = 0 To itemCount - 1
        wordStarts(itemIndex) = VTAdvanceWordOffsetToUtf16Boundary( _
            sourceText, sourceStarts(itemIndex), utf16Cursor, wordCursor)
        wordEnds(itemIndex) = VTAdvanceWordOffsetToUtf16Boundary( _
            sourceText, sourceEnds(itemIndex), utf16Cursor, wordCursor)
    Next itemIndex
End Sub

Private Function VTLatexRedrawWordOffsetsMatch( _
    ByVal documentObject As Document, _
    ByVal targetStart As Long, _
    ByVal targetEnd As Long, _
    ByRef sourceTexts() As String, _
    ByVal itemCount As Long, _
    ByRef candidateStarts() As Long, _
    ByRef candidateEnds() As Long) As Boolean

    Dim itemIndex As Long
    Dim candidateRange As Range

    If documentObject Is Nothing Or itemCount < 1 Then Exit Function
    On Error GoTo NotMatched
    For itemIndex = 0 To itemCount - 1
        If candidateStarts(itemIndex) < 0 Or _
           candidateEnds(itemIndex) <= candidateStarts(itemIndex) Or _
           targetStart + candidateEnds(itemIndex) > targetEnd Then
            Exit Function
        End If
        Set candidateRange = documentObject.Range( _
            Start:=targetStart + candidateStarts(itemIndex), _
            End:=targetStart + candidateEnds(itemIndex))
        If StrComp( _
           candidateRange.Text, sourceTexts(itemIndex), _
           vbBinaryCompare) <> 0 Then
            Exit Function
        End If
    Next itemIndex
    VTLatexRedrawWordOffsetsMatch = True
    Exit Function

NotMatched:
    VTLatexRedrawWordOffsetsMatch = False
End Function

Private Function VTExactLatexTextRangeAt( _
    ByVal documentObject As Document, _
    ByVal candidateStart As Long, _
    ByVal targetEnd As Long, _
    ByVal sourceText As String) As Range

    Dim candidateRange As Range
    Dim utf16Cursor As Long
    Dim wordCursor As Long
    Dim rawLength As Long
    Dim wordLength As Long

    If documentObject Is Nothing Or Len(sourceText) = 0 Or _
       candidateStart < 0 Or candidateStart >= targetEnd Then Exit Function

    rawLength = Len(sourceText)
    If candidateStart + rawLength <= targetEnd Then
        Set candidateRange = documentObject.Range( _
            Start:=candidateStart, End:=candidateStart + rawLength)
        If StrComp(candidateRange.Text, sourceText, vbBinaryCompare) = 0 Then
            Set VTExactLatexTextRangeAt = candidateRange.Duplicate
            Exit Function
        End If
    End If

    utf16Cursor = 0
    wordCursor = 0
    wordLength = VTAdvanceWordOffsetToUtf16Boundary( _
        sourceText, Len(sourceText), utf16Cursor, wordCursor)
    If wordLength > 0 And wordLength <> rawLength And _
       candidateStart + wordLength <= targetEnd Then
        Set candidateRange = documentObject.Range( _
            Start:=candidateStart, End:=candidateStart + wordLength)
        If StrComp(candidateRange.Text, sourceText, vbBinaryCompare) = 0 Then
            Set VTExactLatexTextRangeAt = candidateRange.Duplicate
        End If
    End If
End Function

Private Function VTFindLatexRedrawWordOffsetsByVisibleText( _
    ByVal documentObject As Document, _
    ByVal targetStart As Long, _
    ByVal targetEnd As Long, _
    ByRef sourceTexts() As String, _
    ByVal itemCount As Long, _
    ByRef candidateStarts() As Long, _
    ByRef candidateEnds() As Long) As Boolean

    Const VT_FIND_PREFIX_LENGTH As Long = 180

    Dim searchRange As Range
    Dim exactRange As Range
    Dim searchText As String
    Dim searchPrefix As String
    Dim searchStart As Long
    Dim itemIndex As Long
    Dim foundItem As Boolean

    If documentObject Is Nothing Or itemCount < 1 Or _
       targetEnd <= targetStart Then Exit Function

    ReDim candidateStarts(0 To itemCount - 1)
    ReDim candidateEnds(0 To itemCount - 1)
    searchStart = targetStart

    On Error GoTo NotFound
    For itemIndex = 0 To itemCount - 1
        searchText = sourceTexts(itemIndex)
        If Len(searchText) = 0 Then Exit Function
        searchPrefix = Left$(searchText, VT_FIND_PREFIX_LENGTH)
        ' Word Find treats ^ sequences as special tokens even when wildcard
        ' matching is disabled. LaTeX uses ^ constantly for superscripts, so
        ' double each caret to request a literal caret from Word Find.
        searchPrefix = Replace$(searchPrefix, "^", "^^")
        foundItem = False

        Do While searchStart < targetEnd
            Set searchRange = documentObject.Range( _
                Start:=searchStart, End:=targetEnd)
            With searchRange.Find
                .ClearFormatting
                .Text = searchPrefix
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = True
                .MatchWildcards = False
            End With
            If Not searchRange.Find.Execute Then Exit Do

            Set exactRange = VTExactLatexTextRangeAt( _
                documentObject, searchRange.Start, targetEnd, searchText)
            If Not exactRange Is Nothing Then
                candidateStarts(itemIndex) = exactRange.Start - targetStart
                candidateEnds(itemIndex) = exactRange.End - targetStart
                searchStart = exactRange.End
                foundItem = True
                Exit Do
            End If
            searchStart = searchRange.Start + 1
        Loop
        If Not foundItem Then Exit Function
    Next itemIndex

    VTFindLatexRedrawWordOffsetsByVisibleText = True
    Exit Function

NotFound:
    VTFindLatexRedrawWordOffsetsByVisibleText = False
End Function

Private Sub VTResolveLatexRedrawWordOffsets( _
    ByVal documentObject As Document, _
    ByVal targetStart As Long, _
    ByVal targetEnd As Long, _
    ByVal sourceSnapshot As String, _
    ByRef sourceStarts() As Long, _
    ByRef sourceEnds() As Long, _
    ByRef sourceTexts() As String, _
    ByVal itemCount As Long, _
    ByRef wordStarts() As Long, _
    ByRef wordEnds() As Long)

    Dim itemIndex As Long

    ReDim wordStarts(0 To itemCount - 1)
    ReDim wordEnds(0 To itemCount - 1)
    For itemIndex = 0 To itemCount - 1
        wordStarts(itemIndex) = sourceStarts(itemIndex)
        wordEnds(itemIndex) = sourceEnds(itemIndex)
    Next itemIndex
    If VTLatexRedrawWordOffsetsMatch( _
       documentObject, targetStart, targetEnd, sourceTexts, itemCount, _
       wordStarts, wordEnds) Then
        Exit Sub
    End If

    VTBuildWordOffsetsFromUtf16 _
        sourceSnapshot, sourceStarts, sourceEnds, itemCount, _
        wordStarts, wordEnds
    If VTLatexRedrawWordOffsetsMatch( _
       documentObject, targetStart, targetEnd, sourceTexts, itemCount, _
       wordStarts, wordEnds) Then
        Exit Sub
    End If

    ' Word Range coordinates include hidden Field instructions even though
    ' Range.Text exposes only their visible result. Numbered formulas before a
    ' LaTeX source therefore make parser text offsets diverge from raw Word
    ' positions. Resolve the literal LaTeX spans through Word Find as the final,
    ' exact fallback instead of misreporting that the document text changed.
    If VTFindLatexRedrawWordOffsetsByVisibleText( _
       documentObject, targetStart, targetEnd, sourceTexts, itemCount, _
       wordStarts, wordEnds) Then
        If VTLatexRedrawWordOffsetsMatch( _
           documentObject, targetStart, targetEnd, sourceTexts, itemCount, _
           wordStarts, wordEnds) Then
            Exit Sub
        End If
    End If

    Err.Raise vbObjectError + 7594, "VisualTeX", _
        "Word could not resolve the current LaTeX source ranges. " & _
        "Reopen redraw from the current document."
End Sub

Public Sub VTStartWordLatexRedraw( _
    ByVal redrawScope As String, _
    ByVal outputKind As String)

    Dim sessionId As String
    Dim bookmarkName As String
    Dim sourceDocumentId As String
    Dim sourceText As String
    Dim targetRange As Range
    Dim defaultFontSizePt As Double
    Dim requestJson As String
    Dim operationStage As String
    Dim errorNumber As Long
    Dim errorDescription As String

    If redrawScope <> "selection" And redrawScope <> "document" Then
        VTShowError "Word LaTeX redraw", vbObjectError + 7590, _
            "The LaTeX redraw scope is invalid."
        Exit Sub
    End If
    If outputKind <> "image" And outputKind <> "omml" Then
        VTShowError "Word LaTeX redraw", vbObjectError + 7590, _
            "The LaTeX redraw output kind is invalid."
        Exit Sub
    End If
    On Error GoTo Failed
    operationStage = "require-document"
    VTRequireWritableWordDocument
    operationStage = "refresh-health"
    VTRefreshWordHealthQuietly
    operationStage = "resolve-range"
    If redrawScope = "selection" Then
        Set targetRange = Selection.Range.Duplicate
        If targetRange.Start = targetRange.End Then
            Err.Raise vbObjectError + 7590, "VisualTeX", _
                "Select Word text containing LaTeX before using redraw selection."
        End If
    Else
        Set targetRange = ActiveDocument.Content.Duplicate
    End If
    sourceText = targetRange.Text
    If Len(sourceText) = 0 Then
        Err.Raise vbObjectError + 7590, "VisualTeX", _
            "The selected Word range does not contain text to scan."
    End If
    If VTUtf8ByteLength(sourceText) > _
       VT_WORD_LATEX_REDRAW_MAX_BYTES Then
        Err.Raise vbObjectError + 7590, "VisualTeX", _
            "The LaTeX redraw range exceeds 5 MB. Select a smaller range."
    End If

    operationStage = "create-identities"
    sessionId = VTNewUuidV4()
    bookmarkName = VTDocumentImportBookmarkName(sessionId)
    sourceDocumentId = VTWordDocumentIdentity()
    defaultFontSizePt = VTPreferredWordFormulaFontSize(targetRange)
    If ActiveDocument.Bookmarks.Exists(bookmarkName) Then
        ActiveDocument.Bookmarks(bookmarkName).Delete
    End If
    ActiveDocument.Bookmarks.Add Name:=bookmarkName, Range:=targetRange

    operationStage = "serialize-request"
    requestJson = VTLatexRedrawRequestJson( _
        sessionId, sourceDocumentId, bookmarkName, defaultFontSizePt, _
        redrawScope, outputKind)
    requestJson = VTV2WithDocumentMathFont(requestJson, ActiveDocument)
    operationStage = "write-request"
    VTWriteRequest sessionId, requestJson
    operationStage = "write-source"
    VTWriteLatexRedrawSource sessionId, sourceText
    operationStage = "launch"
    VTLaunchSession VT_WORD_HOST, sessionId
    operationStage = "complete"
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Len(bookmarkName) > 0 Then
        If ActiveDocument.Bookmarks.Exists(bookmarkName) Then
            ActiveDocument.Bookmarks(bookmarkName).Delete
        End If
    End If
    If Len(sessionId) > 0 Then VTDeleteSessionFiles sessionId
    On Error GoTo 0
    VTShowError _
        "Word LaTeX redraw (" & operationStage & ")", _
        errorNumber, errorDescription
End Sub

' A display formula stands in a paragraph of its own, as on Windows: text that shared the
' paragraph with its LaTeX code goes to the paragraph before or after it. Left in the
' paragraph, the equation joined an equation next to it (Word merges adjacent equations)
' or became inline. Table cells keep their paragraph as it is.
Private Sub VTSeparateLatexRedrawDisplayParagraph(ByRef cursorRange As Range)
    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim position As Long

    Set documentObject = cursorRange.Document
    position = cursorRange.Start
    Set paragraphRange = cursorRange.Paragraphs(1).Range.Duplicate
    If paragraphRange.Information(wdWithInTable) Then Exit Sub
    If VTContainsVisibleWordBodyText(documentObject.Range( _
       Start:=position, End:=paragraphRange.End - 1).Text) Then
        documentObject.Range(Start:=position, End:=position).InsertBefore vbCr
    End If
    If VTContainsVisibleWordBodyText(documentObject.Range( _
       Start:=paragraphRange.Start, End:=position).Text) Then
        documentObject.Range(Start:=position, End:=position).InsertBefore vbCr
        position = position + 1
    End If
    Set cursorRange = documentObject.Range(Start:=position, End:=position)
End Sub

Public Function VTContainsVisibleWordBodyText( _
    ByVal value As String) As Boolean

    Dim characterIndex As Long
    Dim currentCharacter As String

    For characterIndex = 1 To Len(value)
        currentCharacter = Mid$(value, characterIndex, 1)
        If currentCharacter <> " " And _
           currentCharacter <> vbTab And _
           currentCharacter <> vbCr And _
           currentCharacter <> vbLf And _
           currentCharacter <> Chr$(7) And _
           currentCharacter <> Chr$(11) And _
           currentCharacter <> Chr$(12) And _
           currentCharacter <> ChrW$(&HA0) Then
            VTContainsVisibleWordBodyText = True
            Exit Function
        End If
    Next characterIndex
End Function

Private Function VTPositionInsideLatexRedrawSource( _
    ByVal absolutePosition As Long, _
    ByVal redrawTargetStart As Long, _
    ByRef redrawSourceStarts() As Long, _
    ByRef redrawSourceEnds() As Long, _
    ByVal redrawSourceCount As Long) As Boolean

    Dim relativePosition As Long
    Dim itemIndex As Long

    relativePosition = absolutePosition - redrawTargetStart
    If relativePosition < 0 Or redrawSourceCount <= 0 Then Exit Function
    For itemIndex = 0 To redrawSourceCount - 1
        If relativePosition < redrawSourceStarts(itemIndex) Then Exit Function
        If relativePosition < redrawSourceEnds(itemIndex) Then
            VTPositionInsideLatexRedrawSource = True
            Exit Function
        End If
    Next itemIndex
End Function

Private Function VTTryResolveNearbyWordFontSize( _
    ByVal documentObject As Document, _
    ByVal startPosition As Long, _
    ByVal minimumPosition As Long, _
    ByVal maximumPosition As Long, _
    ByVal stepValue As Long, _
    ByVal redrawTargetStart As Long, _
    ByRef redrawSourceStarts() As Long, _
    ByRef redrawSourceEnds() As Long, _
    ByVal redrawSourceCount As Long, _
    ByRef fontSizePt As Double) As Boolean

    Const VT_MAX_NEARBY_FONT_PROBES As Long = 256
    Dim documentStart As Long
    Dim documentEnd As Long
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim position As Long
    Dim probeIndex As Long
    Dim probeRange As Range
    Dim probeSize As Double

    fontSizePt = 0#
    If documentObject Is Nothing Then Exit Function
    If stepValue <> -1 And stepValue <> 1 Then Exit Function

    documentStart = documentObject.Content.Start
    documentEnd = documentObject.Content.End - 1
    lowerBound = minimumPosition
    If lowerBound < documentStart Then lowerBound = documentStart
    upperBound = maximumPosition
    If upperBound > documentEnd Then upperBound = documentEnd
    If upperBound < lowerBound Then Exit Function
    If stepValue < 0 And startPosition < lowerBound Then Exit Function
    If stepValue > 0 And startPosition > upperBound Then Exit Function

    position = startPosition
    For probeIndex = 1 To VT_MAX_NEARBY_FONT_PROBES
        If position < lowerBound Or position > upperBound Then Exit For
        Set probeRange = Nothing
        probeSize = 0#
        If Not VTPositionInsideLatexRedrawSource( _
           position, redrawTargetStart, redrawSourceStarts, _
           redrawSourceEnds, redrawSourceCount) Then
            On Error Resume Next
            Set probeRange = documentObject.Range( _
                Start:=position, End:=position + 1)
            If Not probeRange Is Nothing Then
                If VTContainsVisibleWordBodyText(probeRange.Text) Then
                    probeSize = probeRange.Font.Size
                End If
            End If
            Err.Clear
            On Error GoTo 0
        End If
        If VTValidWordFormulaFontSize(probeSize) Then
            fontSizePt = probeSize
            VTTryResolveNearbyWordFontSize = True
            Exit Function
        End If
        position = position + stepValue
    Next probeIndex
End Function

Private Function VTHasVisibleWordTextBesideSource( _
    ByVal sourceRange As Range) As Boolean

    Dim paragraphRange As Range
    Dim surroundingRange As Range
    Dim paragraphBodyEnd As Long

    If sourceRange Is Nothing Then Exit Function
    Set paragraphRange = sourceRange.Paragraphs(1).Range.Duplicate
    paragraphBodyEnd = paragraphRange.End - 1
    If sourceRange.Start > paragraphRange.Start Then
        Set surroundingRange = sourceRange.Document.Range( _
            Start:=paragraphRange.Start, End:=sourceRange.Start)
        If VTContainsVisibleWordBodyText(surroundingRange.Text) Then
            VTHasVisibleWordTextBesideSource = True
            Exit Function
        End If
    End If
    If sourceRange.End < paragraphBodyEnd Then
        Set surroundingRange = sourceRange.Document.Range( _
            Start:=sourceRange.End, End:=paragraphBodyEnd)
        If VTContainsVisibleWordBodyText(surroundingRange.Text) Then
            VTHasVisibleWordTextBesideSource = True
        End If
    End If
End Function

Private Function VTInheritedLatexRedrawFontSize( _
    ByVal sourceRange As Range, _
    ByVal displayMode As String, _
    ByVal redrawTargetStart As Long, _
    ByRef redrawSourceStarts() As Long, _
    ByRef redrawSourceEnds() As Long, _
    ByVal redrawSourceCount As Long) As Double

    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim paragraphStart As Long
    Dim paragraphBodyEnd As Long
    Dim inheritedSize As Double
    Dim styleSize As Double
    Dim sourceSize As Double
    Dim hasVisibleTextBesideSource As Boolean

    If sourceRange Is Nothing Then Exit Function
    VTInheritedLatexRedrawFontSize = VTConversionFontSizeSetting(displayMode)
    If VTValidWordFormulaFontSize(VTInheritedLatexRedrawFontSize) Then Exit Function
    Set documentObject = sourceRange.Document
    Set paragraphRange = sourceRange.Paragraphs(1).Range.Duplicate
    paragraphStart = paragraphRange.Start
    paragraphBodyEnd = paragraphRange.End - 1
    hasVisibleTextBesideSource = _
        VTHasVisibleWordTextBesideSource(sourceRange)

    If displayMode = "inline" Or hasVisibleTextBesideSource Then
        If VTTryResolveNearbyWordFontSize( _
           documentObject, sourceRange.Start - 1, _
           paragraphStart, paragraphBodyEnd, -1, _
           redrawTargetStart, redrawSourceStarts, redrawSourceEnds, _
           redrawSourceCount, inheritedSize) Then
            VTInheritedLatexRedrawFontSize = inheritedSize
            Exit Function
        End If
        If VTTryResolveNearbyWordFontSize( _
           documentObject, sourceRange.End, _
           paragraphStart, paragraphBodyEnd, 1, _
           redrawTargetStart, redrawSourceStarts, redrawSourceEnds, _
           redrawSourceCount, inheritedSize) Then
            VTInheritedLatexRedrawFontSize = inheritedSize
            Exit Function
        End If
    End If

    If VTTryResolveNearbyWordFontSize( _
       documentObject, paragraphStart - 1, _
       documentObject.Content.Start, documentObject.Content.End - 1, _
       -1, redrawTargetStart, redrawSourceStarts, redrawSourceEnds, _
       redrawSourceCount, inheritedSize) Then
        VTInheritedLatexRedrawFontSize = inheritedSize
        Exit Function
    End If
    If VTTryResolveNearbyWordFontSize( _
       documentObject, paragraphRange.End, _
       documentObject.Content.Start, documentObject.Content.End - 1, _
       1, redrawTargetStart, redrawSourceStarts, redrawSourceEnds, _
       redrawSourceCount, inheritedSize) Then
        VTInheritedLatexRedrawFontSize = inheritedSize
        Exit Function
    End If

    On Error Resume Next
    styleSize = paragraphRange.Style.Font.Size
    sourceSize = sourceRange.Font.Size
    On Error GoTo 0
    If VTValidWordFormulaFontSize(styleSize) Then
        VTInheritedLatexRedrawFontSize = styleSize
    ElseIf VTValidWordFormulaFontSize(sourceSize) Then
        VTInheritedLatexRedrawFontSize = sourceSize
    Else
        VTInheritedLatexRedrawFontSize = _
            VTPreferredWordFormulaFontSize(paragraphRange)
    End If
End Function

Public Sub VTResolveWordLatexRedrawFontsDispatch( _
    ByVal sessionId As String, _
    ByVal dispatch As Object)

    Dim manifest As Object
    Dim targetDocument As Document
    Dim anchorBookmark As Bookmark
    Dim targetRange As Range
    Dim sourceRange As Range
    Dim manifestPath As String
    Dim resultPath As String
    Dim sourceDocumentId As String
    Dim bookmarkName As String
    Dim sourceText As String
    Dim sourceSnapshot As String
    Dim displayMode As String
    Dim sourceStarts() As Long
    Dim sourceEnds() As Long
    Dim wordStarts() As Long
    Dim wordEnds() As Long
    Dim sourceTexts() As String
    Dim displayModes() As String
    Dim resultText As String
    Dim targetStart As Long
    Dim targetEnd As Long
    Dim sourceStart As Long
    Dim sourceEnd As Long
    Dim previousEnd As Long
    Dim itemCount As Long
    Dim itemIndex As Long
    Dim fontSizePt As Double

    VTRequireWritableWordDocument
    VTRequireDispatchValue dispatch, "sourceDocumentId"
    VTRequireDispatchValue dispatch, "bookmarkName"
    VTRequireDispatchValue dispatch, "documentImportPath"
    VTRequireDispatchValue dispatch, "fontSizeResultPath"
    sourceDocumentId = CStr(dispatch("sourceDocumentId"))
    bookmarkName = CStr(dispatch("bookmarkName"))
    manifestPath = CStr(dispatch("documentImportPath"))
    resultPath = CStr(dispatch("fontSizeResultPath"))
    VTValidateAbsoluteVisualTeXPath resultPath

    If sourceDocumentId <> VTWordDocumentIdentity() Then
        Err.Raise vbObjectError + 7594, "VisualTeX", _
            "The active Word document changed before LaTeX redraw preflight."
    End If
    If Not ActiveDocument.Bookmarks.Exists(bookmarkName) Then
        Err.Raise vbObjectError + 7594, "VisualTeX", _
            "The LaTeX redraw target range no longer exists."
    End If

    Set manifest = VTReadDocumentImportManifest(manifestPath, sessionId)
    If CStr(manifest("operation")) <> "latexRedraw" Or _
       CStr(manifest("sourceDocumentId")) <> sourceDocumentId Or _
       CStr(manifest("bookmarkName")) <> bookmarkName Or _
       Not IsNumeric(CStr(manifest("itemCount"))) Then
        Err.Raise vbObjectError + 7594, "VisualTeX", _
            "The LaTeX redraw preflight manifest is invalid."
    End If
    itemCount = CLng(manifest("itemCount"))
    If itemCount < 1 Or itemCount > VT_WORD_LATEX_REDRAW_MAX_FORMULAS Then
        Err.Raise vbObjectError + 7594, "VisualTeX", _
            "The LaTeX redraw preflight count is outside the supported range."
    End If

    Set targetDocument = ActiveDocument
    Set anchorBookmark = targetDocument.Bookmarks(bookmarkName)
    Set targetRange = anchorBookmark.Range.Duplicate
    targetStart = targetRange.Start
    targetEnd = targetRange.End
    If targetEnd <= targetStart Then
        Err.Raise vbObjectError + 7594, "VisualTeX", _
            "The LaTeX redraw target range is empty."
    End If
    sourceSnapshot = VTReadText( _
        VTSessionDirectory(sessionId) & VT_WORD_LATEX_REDRAW_SOURCE_FILE, _
        VT_WORD_LATEX_REDRAW_MAX_BYTES)
    If Len(sourceSnapshot) = 0 Then
        Err.Raise vbObjectError + 7594, "VisualTeX", _
            "The LaTeX redraw source snapshot is missing."
    End If

    ReDim sourceStarts(0 To itemCount - 1)
    ReDim sourceEnds(0 To itemCount - 1)
    ReDim sourceTexts(0 To itemCount - 1)
    ReDim displayModes(0 To itemCount - 1)
    previousEnd = 0
    For itemIndex = 0 To itemCount - 1
        sourceStarts(itemIndex) = VTDocumentImportSourceOffset( _
            manifest, itemIndex, "sourceStart")
        sourceEnds(itemIndex) = VTDocumentImportSourceOffset( _
            manifest, itemIndex, "sourceEnd")
        If sourceStarts(itemIndex) < previousEnd Or _
           sourceEnds(itemIndex) <= sourceStarts(itemIndex) Or _
           sourceEnds(itemIndex) > Len(sourceSnapshot) Then
            Err.Raise vbObjectError + 7594, "VisualTeX", _
                "LaTeX redraw preflight ranges overlap or exceed the source snapshot."
        End If
        sourceTexts(itemIndex) = VTBase64UrlDecodeUtf8( _
            VTDocumentImportRequired( _
                manifest, itemIndex, "sourceTextBase64"))
        If StrComp( _
           Mid$(sourceSnapshot, sourceStarts(itemIndex) + 1, _
                sourceEnds(itemIndex) - sourceStarts(itemIndex)), _
           sourceTexts(itemIndex), vbBinaryCompare) <> 0 Then
            Err.Raise vbObjectError + 7594, "VisualTeX", _
                "The LaTeX redraw manifest no longer matches its source snapshot."
        End If
        displayModes(itemIndex) = VTDocumentImportRequired( _
            manifest, itemIndex, "displayMode")
        If displayModes(itemIndex) <> "inline" And _
           displayModes(itemIndex) <> "block" Then
            Err.Raise vbObjectError + 7594, "VisualTeX", _
                "A LaTeX redraw preflight display mode is invalid."
        End If
        previousEnd = sourceEnds(itemIndex)
    Next itemIndex

    VTResolveLatexRedrawWordOffsets _
        targetDocument, targetStart, targetEnd, sourceSnapshot, _
        sourceStarts, sourceEnds, sourceTexts, itemCount, _
        wordStarts, wordEnds

    resultText = "itemCount=" & CStr(itemCount) & vbLf
    For itemIndex = 0 To itemCount - 1
        If wordEnds(itemIndex) > targetEnd - targetStart Then
            Err.Raise vbObjectError + 7594, "VisualTeX", _
                "A LaTeX redraw range exceeds the current Word target."
        End If
        Set sourceRange = targetDocument.Range( _
            Start:=targetStart + wordStarts(itemIndex), _
            End:=targetStart + wordEnds(itemIndex))
        fontSizePt = VTInheritedLatexRedrawFontSize( _
            sourceRange, displayModes(itemIndex), targetStart, _
            wordStarts, wordEnds, itemCount)
        If Not VTValidWordFormulaFontSize(fontSizePt) Then
            Err.Raise vbObjectError + 7594, "VisualTeX", _
                "A LaTeX redraw source formula has an invalid Word font size."
        End If
        resultText = resultText & _
            "item" & CStr(itemIndex) & "fontSizePt=" & _
            VTJsonNumber(fontSizePt) & vbLf
    Next itemIndex

    VTWriteTextAtomic resultPath, resultText
End Sub

Private Sub VTResolveFormulaRestoreWordOffsets( _
    ByVal documentObject As Document, _
    ByVal targetStart As Long, _
    ByVal targetEnd As Long, _
    ByVal manifest As Object, _
    ByVal sourceKind As String, _
    ByVal itemCount As Long, _
    ByRef wordStarts() As Long, _
    ByRef wordEnds() As Long)

    Dim targetRange As Range
    Dim sourceRange As Range
    Dim formulaShape As InlineShape
    Dim formulaMath As OMath
    Dim nativeBookmark As Bookmark
    Dim targets As Collection
    Dim expectedFormulaId As String
    Dim actualFormulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean
    Dim itemIndex As Long
    Dim previousEnd As Long

    If documentObject Is Nothing Or targetEnd <= targetStart Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The formula restore target is no longer available."
    End If
    If sourceKind <> "image" And sourceKind <> "omml" Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The formula restore source kind is invalid."
    End If

    Set targetRange = documentObject.Range( _
        Start:=targetStart, End:=targetEnd)
    Set targets = VTCollectFormulaRestoreTargets(targetRange, sourceKind)
    If targets.Count <> itemCount Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The Word formula inventory changed while format conversion was running."
    End If

    ReDim wordStarts(0 To itemCount - 1)
    ReDim wordEnds(0 To itemCount - 1)
    previousEnd = 0
    For itemIndex = 0 To itemCount - 1
        expectedFormulaId = VTDispatchOptional( _
            manifest, VTDocumentImportItemKey(itemIndex, "formulaId"))
        Set sourceRange = Nothing
        If sourceKind = "image" Then
            Set formulaShape = targets(itemIndex + 1)
            actualFormulaId = ""
            displayMode = ""
            numbered = False
            encodedMetadata = ""
            metadataNeedsWrite = False
            formatNeedsWrite = False
            If Not VTTryResolveVisualTeXInlineShapeReference( _
               formulaShape, actualFormulaId, displayMode, numbered, _
               encodedMetadata, metadataNeedsWrite, formatNeedsWrite) Or _
               (Len(expectedFormulaId) > 0 And _
                StrComp(actualFormulaId, expectedFormulaId, _
                    vbTextCompare) <> 0) Then
                Err.Raise vbObjectError + 7591, "VisualTeX", _
                    "A Word image formula identity changed during format conversion."
            End If
            Set sourceRange = VTVisualTeXImageContainerRange(formulaShape)
        Else
            Set formulaMath = targets(itemIndex + 1)
            Set sourceRange = formulaMath.Range.Duplicate
            If VTDispatchOptional( _
               manifest, VTDocumentImportItemKey( _
                   itemIndex, "preserveIdentity")) = "1" Then
                Set nativeBookmark = Nothing
                ' Word can move the edge of VT_F_ while an equation/REF field
                ' is rebuilt. Range.Bookmarks only returns contained bookmarks;
                ' resolve the known identity directly, then compare live OMaths.
                ' Keep object checks separate: VBA And/Or never short-circuit.
                If Not documentObject.Bookmarks.Exists( _
                   VTNativeFormulaBookmarkName(expectedFormulaId)) Then
                    Err.Raise vbObjectError + 7591, "VisualTeX", _
                        "A managed Word OMML formula bookmark disappeared during format conversion."
                End If
                Set nativeBookmark = documentObject.Bookmarks( _
                    VTNativeFormulaBookmarkName(expectedFormulaId))
                Set formulaMath = VTNativeMathForBookmark(nativeBookmark)
                If formulaMath Is Nothing Then
                    Err.Raise vbObjectError + 7591, "VisualTeX", _
                        "A managed Word OMML formula lost its equation during format conversion."
                End If
                If formulaMath.Range.Start <> sourceRange.Start Or _
                   formulaMath.Range.End <> sourceRange.End Then
                    Err.Raise vbObjectError + 7591, "VisualTeX", _
                        "A managed Word OMML formula identity changed during format conversion."
                End If
            End If
        End If
        If sourceRange Is Nothing Then
            Err.Raise vbObjectError + 7591, "VisualTeX", _
                "Word could not re-resolve a formula during format conversion."
        End If
        wordStarts(itemIndex) = sourceRange.Start - targetStart
        wordEnds(itemIndex) = sourceRange.End - targetStart
        If wordStarts(itemIndex) < previousEnd Or _
           wordEnds(itemIndex) <= wordStarts(itemIndex) Or _
           wordEnds(itemIndex) > targetEnd - targetStart Then
            Err.Raise vbObjectError + 7591, "VisualTeX", _
                "The live Word formula ranges overlap or exceed the conversion target."
        End If
        previousEnd = wordEnds(itemIndex)
    Next itemIndex
End Sub

Public Sub VTCommitWordLatexRedrawDispatch( _
    ByVal sessionId As String, _
    ByVal manifest As Object, _
    ByVal sourceDocumentId As String, _
    ByVal bookmarkName As String, _
    ByVal outputKind As String, _
    ByVal itemCount As Long)

    Dim targetDocument As Document
    Dim anchorBookmark As Bookmark
    Dim targetRange As Range
    Dim sourceRange As Range
    Dim cursorRange As Range
    Dim rollbackRange As Range
    Dim sharedVectorDocument As Document
    Dim sharedNativeDocument As Document
    Dim undoRecord As Object
    Dim insertedFormulaIds As Collection
    Dim numberedFormulaIds As Collection
    Dim numberedRestoreRanges As Collection
    Dim sourceStarts() As Long
    Dim sourceEnds() As Long
    Dim wordStarts() As Long
    Dim wordEnds() As Long
    Dim sourceTexts() As String
    Dim sourceFontSizes() As Double
    Dim originalSource As String
    Dim sourceSnapshot As String
    Dim sourceText As String
    Dim formulaId As String
    Dim itemKind As String
    Dim operationName As String
    Dim sourceKind As String
    Dim replacementText As String
    Dim formulaFailureDescription As String
    Dim sharedVectorDocumentPath As String
    Dim nativeBatchDocumentPath As String
    Dim targetStart As Long
    Dim targetEnd As Long
    Dim originalTargetLength As Long
    Dim previousEnd As Long
    Dim itemIndex As Long
    Dim completedItems As Long
    Dim nativeBatchFormulaCount As Long
    Dim numberedFormulaCount As Long
    Dim progressInterval As Long
    Dim previousDisplayAlerts As Long
    Dim previousScreenUpdating As Boolean
    Dim displayAlertsChanged As Boolean
    Dim screenUpdatingChanged As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim transactionStage As String
    Dim internalMutationStarted As Boolean
    Dim undoRecordStarted As Boolean
    Dim mutationOccurred As Boolean
    Dim rollbackVerified As Boolean
    Dim atomicFormulaConversion As Boolean
    Dim deferNumberingReconcile As Boolean
    Dim numberParagraphCleared As Boolean
    Dim replacementStart As Long

    On Error GoTo Failed
    transactionStage = "redraw-validate-target"
    operationName = CStr(manifest("operation"))
    atomicFormulaConversion = _
        (operationName = "formulaRestore" And outputKind <> "latex")
    If operationName <> "latexRedraw" And _
       operationName <> "formulaRestore" Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The Word range replacement operation is invalid."
    End If
    If itemCount < 1 Or itemCount > VT_WORD_LATEX_REDRAW_MAX_FORMULAS Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The LaTeX redraw formula count is outside the supported range."
    End If
    If sourceDocumentId <> VTWordDocumentIdentity() Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The active Word document changed while LaTeX redraw was open."
    End If
    Set targetDocument = ActiveDocument
    If Not targetDocument.Bookmarks.Exists(bookmarkName) Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The LaTeX redraw target range no longer exists."
    End If
    Set anchorBookmark = targetDocument.Bookmarks(bookmarkName)
    Set targetRange = anchorBookmark.Range.Duplicate
    targetStart = targetRange.Start
    targetEnd = targetRange.End
    originalTargetLength = targetEnd - targetStart
    originalSource = targetRange.Text
    If targetEnd <= targetStart Then
        Err.Raise vbObjectError + 7591, "VisualTeX", _
            "The LaTeX redraw target range is empty."
    End If
    If operationName = "latexRedraw" Then
        sourceSnapshot = VTReadText( _
            VTSessionDirectory(sessionId) & VT_WORD_LATEX_REDRAW_SOURCE_FILE, _
            VT_WORD_LATEX_REDRAW_MAX_BYTES)
        If Len(sourceSnapshot) = 0 Then
            Err.Raise vbObjectError + 7591, "VisualTeX", _
                "The LaTeX redraw source snapshot is missing."
        End If
    Else
        sourceSnapshot = originalSource
    End If

    ReDim sourceStarts(0 To itemCount - 1)
    ReDim sourceEnds(0 To itemCount - 1)
    ReDim sourceTexts(0 To itemCount - 1)
    ReDim sourceFontSizes(0 To itemCount - 1)
    previousEnd = 0
    For itemIndex = 0 To itemCount - 1
        transactionStage = "redraw-preflight item " & CStr(itemIndex)
        itemKind = VTDocumentImportRequired(manifest, itemIndex, "kind")
        If itemKind = "formula" Then
            formulaId = VTDocumentImportRequired( _
                manifest, itemIndex, "formulaId")
            If Not VTIsCanonicalUuid(formulaId) Then
                Err.Raise vbObjectError + 7591, "VisualTeX", _
                    "A Word replacement formula id is invalid."
            End If
            If VTDocumentImportRequired( _
               manifest, itemIndex, "numbered") = "1" Then
                numberedFormulaCount = numberedFormulaCount + 1
            End If
        ElseIf itemKind = "text" Then
            If operationName <> "formulaRestore" Or _
               outputKind <> "latex" Then
                Err.Raise vbObjectError + 7591, "VisualTeX", _
                    "Only formula restore can write LaTeX text items."
            End If
            replacementText = VTBase64UrlDecodeUtf8( _
                VTDocumentImportRequired( _
                    manifest, itemIndex, "textBase64"))
            If Len(replacementText) = 0 Then
                Err.Raise vbObjectError + 7591, "VisualTeX", _
                    "A restored LaTeX text item is empty."
            End If
        Else
            Err.Raise vbObjectError + 7591, "VisualTeX", _
                "The Word range replacement item kind is invalid."
        End If
        sourceStarts(itemIndex) = VTDocumentImportSourceOffset( _
            manifest, itemIndex, "sourceStart")
        sourceEnds(itemIndex) = VTDocumentImportSourceOffset( _
            manifest, itemIndex, "sourceEnd")
        If sourceStarts(itemIndex) < previousEnd Or _
           sourceEnds(itemIndex) <= sourceStarts(itemIndex) Then
            Err.Raise vbObjectError + 7591, "VisualTeX", _
                "LaTeX redraw source ranges overlap or are invalid."
        End If
        If operationName = "latexRedraw" Then
            If sourceEnds(itemIndex) > Len(sourceSnapshot) Then
                Err.Raise vbObjectError + 7591, "VisualTeX", _
                    "A LaTeX redraw source range exceeds the source snapshot."
            End If
        ElseIf sourceEnds(itemIndex) > originalTargetLength Then
            Err.Raise vbObjectError + 7591, "VisualTeX", _
                "A formula restore source range exceeds the Word target."
        End If
        sourceTexts(itemIndex) = VTBase64UrlDecodeUtf8( _
            VTDocumentImportRequired( _
                manifest, itemIndex, "sourceTextBase64"))
        If operationName = "latexRedraw" Then
            If StrComp( _
               Mid$(sourceSnapshot, sourceStarts(itemIndex) + 1, _
                    sourceEnds(itemIndex) - sourceStarts(itemIndex)), _
               sourceTexts(itemIndex), vbBinaryCompare) <> 0 Then
                Err.Raise vbObjectError + 7591, "VisualTeX", _
                    "The LaTeX redraw manifest no longer matches its source snapshot."
            End If
        End If
        If itemKind = "formula" Then
            sourceFontSizes(itemIndex) = VTDispatchPositiveDouble( _
                manifest, VTDocumentImportItemKey(itemIndex, "fontSizePt"))
            If Not VTValidWordFormulaFontSize(sourceFontSizes(itemIndex)) Then
                Err.Raise vbObjectError + 7591, "VisualTeX", _
                    "A Word replacement formula has an invalid inherited font size."
            End If
        Else
            sourceFontSizes(itemIndex) = 0#
        End If
        previousEnd = sourceEnds(itemIndex)
    Next itemIndex

    ' A document format conversion keeps every existing VT_N_ helper in place,
    ' so rebuilding each visible image number does not change sequence order.
    ' Defer the expensive full-document replay until every replacement is
    ' complete, then validate the complete mixed Word/VisualTeX sequence once.
    deferNumberingReconcile = _
        atomicFormulaConversion And outputKind = "image" And _
        numberedFormulaCount > 1

    If operationName = "latexRedraw" Then
        VTResolveLatexRedrawWordOffsets _
            targetDocument, targetStart, targetEnd, sourceSnapshot, _
            sourceStarts, sourceEnds, sourceTexts, itemCount, _
            wordStarts, wordEnds
    Else
        transactionStage = "redraw-resolve-live-formulas"
        sourceKind = VTDispatchOptional(manifest, "sourceKind")
        VTResolveFormulaRestoreWordOffsets _
            targetDocument, targetStart, targetEnd, manifest, _
            sourceKind, itemCount, wordStarts, wordEnds
    End If

    For itemIndex = 0 To itemCount - 1
        transactionStage = "redraw-live-range item " & CStr(itemIndex)
        If wordEnds(itemIndex) > originalTargetLength Then
            Err.Raise vbObjectError + 7591, "VisualTeX", _
                "A Word replacement range exceeds the current target."
        End If
        Set sourceRange = targetDocument.Range( _
            Start:=targetStart + wordStarts(itemIndex), _
            End:=targetStart + wordEnds(itemIndex))
        If operationName = "latexRedraw" And StrComp( _
           sourceRange.Text, sourceTexts(itemIndex), _
           vbBinaryCompare) <> 0 Then
            Err.Raise vbObjectError + 7591, "VisualTeX", _
                "Word text changed after LaTeX redraw opened. " & _
                "Please reopen redraw from the current selection."
        End If
    Next itemIndex

    VTWordPerformanceMark "redraw-live-ranges-ready"
    transactionStage = "redraw-disable-ui"
    On Error Resume Next
    Err.Clear
    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    screenUpdatingChanged = (Err.Number = 0)
    Err.Clear
    previousDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = wdAlertsNone
    displayAlertsChanged = (Err.Number = 0)
    Err.Clear
    On Error GoTo Failed

    If outputKind = "image" Then
        sharedVectorDocumentPath = VTDocumentImportRequired( _
            manifest, 0, "vectorDocumentPath")
        If CLng(Val(VTDispatchOptional( _
           manifest, VTDocumentImportItemKey(0, "vectorDocumentIndex")))) > 0 Then
            transactionStage = "redraw-open-shared-vector-document"
            VTValidateAbsoluteVisualTeXPath sharedVectorDocumentPath
            If Not VTPathFileExists(sharedVectorDocumentPath) Then
                Err.Raise vbObjectError + 7591, "VisualTeX", _
                    "The shared LaTeX redraw vector document is missing."
            End If
            Set sharedVectorDocument = Documents.Open( _
                FileName:=sharedVectorDocumentPath, _
                ConfirmConversions:=False, _
                ReadOnly:=True, _
                AddToRecentFiles:=False, _
                Visible:=False)
            If sharedVectorDocument.InlineShapes.Count <> itemCount Then
                Err.Raise vbObjectError + 7591, "VisualTeX", _
                    "The shared LaTeX redraw vector document is incomplete."
            End If
            targetDocument.Activate
        End If
    ElseIf outputKind = "omml" Then
        nativeBatchDocumentPath = VTDispatchOptional( _
            manifest, "nativeBatchDocumentPath")
        If Len(nativeBatchDocumentPath) > 0 Then
            If Not IsNumeric(VTDispatchOptional( _
               manifest, "nativeBatchFormulaCount")) Then
                Err.Raise vbObjectError + 7591, "VisualTeX", _
                    "The shared LaTeX redraw native batch count is invalid."
            End If
            nativeBatchFormulaCount = CLng(VTDispatchOptional( _
                manifest, "nativeBatchFormulaCount"))
            If nativeBatchFormulaCount <> itemCount Then
                Err.Raise vbObjectError + 7591, "VisualTeX", _
                    "The shared LaTeX redraw native batch count is incomplete."
            End If
            transactionStage = "redraw-open-shared-native-document"
            VTValidateAbsoluteVisualTeXPath nativeBatchDocumentPath
            If Not VTPathFileExists(nativeBatchDocumentPath) Then
                Err.Raise vbObjectError + 7591, "VisualTeX", _
                    "The shared LaTeX redraw native document is missing."
            End If
            Set sharedNativeDocument = Documents.Open( _
                FileName:=nativeBatchDocumentPath, _
                ConfirmConversions:=False, _
                ReadOnly:=True, _
                AddToRecentFiles:=False, _
                Visible:=False)
            If sharedNativeDocument.OMaths.Count <> itemCount Then
                Err.Raise vbObjectError + 7591, "VisualTeX", _
                    "The shared LaTeX redraw native document is incomplete."
            End If
            targetDocument.Activate
        End If
    End If

    VTWordPerformanceMark "redraw-staging-documents-open"
    ' Remove the temporary anchor before the Undo record starts: deleted
    ' inside it, an Undo of the redraw would put the VT_D_ Bookmark back into
    ' the user's document for good.
    transactionStage = "redraw-remove-anchor"
    anchorBookmark.Delete
    transactionStage = "redraw-start-undo"
    Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
    If operationName = "formulaRestore" Then
        undoRecord.StartCustomRecord "VisualTeX Restore Word Formulas"
    Else
        undoRecord.StartCustomRecord "VisualTeX Redraw LaTeX Formulas"
    End If
    undoRecordStarted = True
    Set insertedFormulaIds = New Collection
    Set numberedFormulaIds = New Collection
    Set numberedRestoreRanges = New Collection
    VTBeginWordInternalMutation
    internalMutationStarted = True
    VTUpdateDocumentImportProgress sessionId, 0, itemCount, "inserting"
    progressInterval = itemCount \ 10
    If progressInterval < 1 Then progressInterval = 1

    For itemIndex = itemCount - 1 To 0 Step -1
        transactionStage = "redraw-replace item " & CStr(itemIndex)
        Set sourceRange = targetDocument.Range( _
            Start:=targetStart + wordStarts(itemIndex), _
            End:=targetStart + wordEnds(itemIndex))
        itemKind = VTDocumentImportRequired(manifest, itemIndex, "kind")
        If itemKind = "text" Then
            replacementText = VTBase64UrlDecodeUtf8( _
                VTDocumentImportRequired( _
                    manifest, itemIndex, "textBase64"))
            Set cursorRange = VTReplaceFormulaRestoreRangeWithLatex( _
                sourceRange, replacementText, mutationOccurred, _
                numberedFormulaIds, numberedRestoreRanges)
            cursorRange.Collapse wdCollapseEnd
        Else
            numberParagraphCleared = False
            If sourceKind = "image" Then
                Set sourceRange = VTNumberedImageReplacementRange( _
                    sourceRange, numberParagraphCleared)
            End If
            replacementStart = sourceRange.Start
            sourceRange.Delete
            mutationOccurred = True
            Set cursorRange = targetDocument.Range( _
                Start:=replacementStart, End:=replacementStart)
            If numberParagraphCleared Then
                VTNormalizePlainWordParagraph _
                    cursorRange.Paragraphs(1).Range.Duplicate
            End If
            If operationName = "latexRedraw" And _
               VTDispatchOptional(manifest, VTDocumentImportItemKey( _
                   itemIndex, "displayMode")) = "block" Then
                VTSeparateLatexRedrawDisplayParagraph cursorRange
            End If
            formulaFailureDescription = ""
            If Not sharedVectorDocument Is Nothing Then
                If Not VTTryDocumentImportInsertFormula( _
                   cursorRange, manifest, itemIndex, outputKind, _
                   insertedFormulaIds, formulaFailureDescription, _
                   sourceFontSizes(itemIndex), True, _
                   sharedVectorDocument, Empty, _
                   atomicFormulaConversion, _
                   deferNumberingReconcile) Then
                    Err.Raise vbObjectError + 7591, "VisualTeX", _
                        "Formula replacement failed: " & formulaFailureDescription
                End If
            ElseIf Not sharedNativeDocument Is Nothing Then
                If Not VTTryDocumentImportInsertFormula( _
                   cursorRange, manifest, itemIndex, outputKind, _
                   insertedFormulaIds, formulaFailureDescription, _
                   sourceFontSizes(itemIndex), True, Empty, _
                   sharedNativeDocument, atomicFormulaConversion, _
                   deferNumberingReconcile) Then
                    Err.Raise vbObjectError + 7591, "VisualTeX", _
                        "Formula replacement failed: " & formulaFailureDescription
                End If
            Else
                If Not VTTryDocumentImportInsertFormula( _
                   cursorRange, manifest, itemIndex, outputKind, _
                   insertedFormulaIds, formulaFailureDescription, _
                   sourceFontSizes(itemIndex), True, Empty, Empty, _
                   atomicFormulaConversion, _
                   deferNumberingReconcile) Then
                    Err.Raise vbObjectError + 7591, "VisualTeX", _
                        "Formula replacement failed: " & formulaFailureDescription
                End If
            End If
        End If
        VTWordPerformanceMark "redraw-item-" & CStr(itemIndex) & "-complete"
        completedItems = itemCount - itemIndex
        If completedItems = itemCount Or _
           completedItems Mod progressInterval = 0 Then
            VTUpdateDocumentImportProgress _
                sessionId, completedItems, itemCount, "inserting"
        End If
    Next itemIndex

    ' Each image number was written while the formulas before it were still
    ' unconverted, so its chapter restart could not see them.
    If atomicFormulaConversion And outputKind = "image" And _
       numberedFormulaCount > 0 Then
        transactionStage = "redraw-image-number-restarts"
        VTRepairImageNumberRestartsFrom targetDocument, targetStart
    End If
    If deferNumberingReconcile Then
        transactionStage = "redraw-reconcile-numbering"
        VTReconcileEquationNumbers targetDocument, -1, True
    End If
    VTWordPerformanceMark "redraw-numbering-complete"

    If operationName = "formulaRestore" And outputKind = "latex" Then
        VTFinalizeFormulaRestoreNumbering _
            targetDocument, numberedFormulaIds, numberedRestoreRanges
    End If

    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    If Not sharedVectorDocument Is Nothing Then
        sharedVectorDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set sharedVectorDocument = Nothing
        targetDocument.Activate
    End If
    If Not sharedNativeDocument Is Nothing Then
        sharedNativeDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set sharedNativeDocument = Nothing
        targetDocument.Activate
    End If
    If undoRecordStarted Then
        undoRecord.EndCustomRecord
        undoRecordStarted = False
    End If
    If displayAlertsChanged Then
        Application.DisplayAlerts = previousDisplayAlerts
        displayAlertsChanged = False
    End If
    If screenUpdatingChanged Then
        Application.ScreenUpdating = previousScreenUpdating
        screenUpdatingChanged = False
    End If
    If Not cursorRange Is Nothing Then cursorRange.Select
    VTWordPerformanceMark "redraw-cleanup-complete"
    VTUpdateDocumentImportProgress _
        sessionId, itemCount, itemCount, "complete"
    Application.StatusBar = ""
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    VTUpdateDocumentImportProgress sessionId, 0, itemCount, "error"
    VTWriteWordFailureTrace _
        sessionId, transactionStage, errorNumber, errorDescription
    If Not sharedVectorDocument Is Nothing Then
        sharedVectorDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set sharedVectorDocument = Nothing
        targetDocument.Activate
    End If
    If Not sharedNativeDocument Is Nothing Then
        sharedNativeDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set sharedNativeDocument = Nothing
        targetDocument.Activate
    End If
    If undoRecordStarted Then
        undoRecord.EndCustomRecord
        undoRecordStarted = False
    End If
    rollbackVerified = Not mutationOccurred
    If mutationOccurred Then
        targetDocument.Undo 1
        Err.Clear
        Set rollbackRange = targetDocument.Range( _
            Start:=targetStart, End:=targetStart + originalTargetLength)
        rollbackVerified = _
            StrComp(rollbackRange.Text, originalSource, vbBinaryCompare) = 0
    End If
    If Not targetDocument Is Nothing Then
        For itemIndex = 0 To itemCount - 1
            itemKind = VTDocumentImportRequired( _
                manifest, itemIndex, "kind")
            If itemKind = "formula" Then
                formulaId = VTDocumentImportRequired( _
                    manifest, itemIndex, "formulaId")
                If VTIsCanonicalUuid(formulaId) And _
                   VTDispatchOptional( _
                       manifest, _
                       VTDocumentImportItemKey( _
                           itemIndex, "preserveIdentity")) <> "1" Then
                    VTDeleteDocumentImportedFormulaState _
                        targetDocument, formulaId
                End If
            End If
        Next itemIndex
        If rollbackVerified And _
           Not targetDocument.Bookmarks.Exists(bookmarkName) Then
            Set rollbackRange = targetDocument.Range( _
                Start:=targetStart, End:=targetStart + originalTargetLength)
            targetDocument.Bookmarks.Add _
                Name:=bookmarkName, Range:=rollbackRange
        End If
    End If
    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    If displayAlertsChanged Then
        Application.DisplayAlerts = previousDisplayAlerts
        displayAlertsChanged = False
    End If
    If screenUpdatingChanged Then
        Application.ScreenUpdating = previousScreenUpdating
        screenUpdatingChanged = False
    End If
    On Error GoTo 0
    If Not rollbackVerified Then
        Err.Raise vbObjectError + 7592, _
            "VisualTeX Word LaTeX redraw", _
            transactionStage & ": " & errorDescription & _
            " Word could not verify the automatic rollback."
    End If
    Err.Raise errorNumber, "VisualTeX Word LaTeX redraw", _
        transactionStage & ": " & errorDescription
End Sub
