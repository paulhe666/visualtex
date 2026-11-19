Attribute VB_Name = "VTWordConversion"
Option Explicit
Option Private Module

' Format conversion between image formulas and Word equations, conversion font sizes, and the upgrades of
' older VisualTeX forms (Update Equation Format, legacy layouts).

Private Const VT_WORD_DISPLAY_CONVERSION_FONT_CONTROL_ID As String = _
    "VisualTeX.Mac.Word.DisplayConversionFontSize"
' The global sizes conversions give inline and display formulas. 0 keeps the
' size the conversion would otherwise use (body text or the formula's own).
Private VT_WORD_CONVERSION_FONT_SIZES_LOADED As Boolean
Private VT_WORD_INLINE_CONVERSION_FONT_SIZE_PT As Double
Private VT_WORD_DISPLAY_CONVERSION_FONT_SIZE_PT As Double

Public Function VTMigrateLegacyImageEquationSequenceLayouts( _
    ByVal documentObject As Document) As Long

    Dim candidateField As Field
    Dim visibleNumberField As Field
    Dim visibleNumberRange As Range
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

    ' Migrate direct SEQ and plain-text mirrors. Preserve the current layout:
    ' an external SEQ helper with a visible REF.
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
                                Set visibleNumberField = Nothing
                                Set visibleNumberRange = Nothing
                                If VTHelperParagraphOwnsNativeEquationSequence( _
                                   sequenceParagraph) Then
                                    Set visibleNumberField = _
                                        VTImageEquationReferenceField( _
                                            formulaRange, formulaId)
                                    If Not visibleNumberField Is Nothing Then
                                        Set visibleNumberRange = _
                                            VTImageEquationNumberRange( _
                                                formulaRange, visibleNumberField)
                                    End If
                                End If
                                If visibleNumberRange Is Nothing Then
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

    VTWordPerformanceNote "imageNumberMigrations", CStr(migrationIds.Count)
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

Public Sub VTMigrateDocumentImageMacroButtons( _
    ByVal documentObject As Document)

    Dim formulaShape As InlineShape
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim imageIndex As Long
    Dim wasSaved As Boolean
    Dim hasManagedImage As Boolean
    Dim inkCenterYRatio As Double
    Dim hasNumberedImage As Boolean
    Dim requiresNumberReconcile As Boolean
    Dim internalMutationStarted As Boolean
    Dim migrationErrorNumber As Long
    Dim migrationErrorDescription As String

    If documentObject Is Nothing Then Exit Sub
    If documentObject.ReadOnly Or _
       documentObject.ProtectionType <> wdNoProtection Then Exit Sub
    If VTDocumentImageMacroSchemaCurrent(documentObject) Then Exit Sub
    For Each formulaShape In documentObject.InlineShapes
        If VTIsVisualTeXInlineShape(formulaShape) Then hasManagedImage = True
    Next formulaShape
    If Not hasManagedImage Then Exit Sub

    On Error GoTo MigrationFailed
    wasSaved = documentObject.Saved
    VTBeginWordInternalMutation
    internalMutationStarted = True
    VTCleanupDocumentEmptyVisualTeXImageMacroButtons documentObject
    ' Schema 10 keeps the hidden increment and visible current-value field in
    ' the same paragraph as the image, ends the native MACROBUTTON label at its
    ' real closing parenthesis and advances the private VisualTeXEquation stream
    ' instead of Word's built-in Equation stream. Existing helpers and Schema-9
    ' PlaceRefs are retired only after their reference targets have been moved.
    VTMigrateLegacyNumberedImageTables documentObject
    For imageIndex = documentObject.InlineShapes.Count To 1 Step -1
        Set formulaShape = documentObject.InlineShapes(imageIndex)
        If VTIsVisualTeXInlineShape(formulaShape) Then
            Set formulaShape = _
                VTEnsureVisualTeXImageMacroButton(formulaShape)
            If VTTryParseFormulaReference( _
               formulaShape.Title, formulaId, displayMode, numbered) Then
                If displayMode = "block" Then
                    If Not VTTryReadWordImageInkCenterYRatio( _
                       documentObject, formulaId, inkCenterYRatio) Then
                        If Not VTTryReadCachedWordImageInkCenterYRatio( _
                           formulaId, inkCenterYRatio) Then
                            inkCenterYRatio = 0.5
                        End If
                        VTSetWordImageInkCenterYRatio _
                            documentObject, formulaId, inkCenterYRatio
                    End If
                    If numbered Then
                        hasNumberedImage = True
                    Else
                        VTNormalizeUnnumberedDisplayParagraph formulaShape.Range
                        VTNormalizeImageDisplayParagraph formulaShape.Range
                    End If
                End If
            End If
        End If
    Next imageIndex
    ' The numbers in the current layout, in the document's own numbering
    ' style. A document VisualTeX wrote has them already (the schema marker
    ' is only written here, so every such document comes this way when it is
    ' opened); only numbers in an older layout are written again.
    If hasNumberedImage Then
        If VTUpgradeImageNumberFormats(documentObject, True) > 0 Then _
            requiresNumberReconcile = True
    End If
    If requiresNumberReconcile Then
        VTReconcileEquationNumbers documentObject
    End If
    VTSetDocumentVariable _
        documentObject, VT_WORD_IMAGE_MACRO_SCHEMA_VARIABLE, _
        VT_WORD_IMAGE_MACRO_SCHEMA_VERSION
    If wasSaved Then documentObject.Saved = True
    VTEndWordInternalMutation
    Exit Sub

MigrationFailed:
    migrationErrorNumber = Err.Number
    migrationErrorDescription = Err.Description
    If internalMutationStarted Then VTEndWordInternalMutation
    On Error Resume Next
    If wasSaved Then documentObject.Saved = True
    On Error GoTo 0
    Err.Raise migrationErrorNumber, "VisualTeX image migration", _
        migrationErrorDescription
End Sub

Public Sub VTReadWordImageConversionState( _
    ByVal formulaShape As InlineShape, _
    ByVal formulaId As String, _
    ByRef fontSizePt As Double, _
    ByRef referenceWidthPt As Double, _
    ByRef referenceHeightPt As Double, _
    ByRef referenceBaselinePt As Double, _
    ByRef observedWordFontSizePt As Double)

    Dim stateReadable As Boolean
    Dim stateErrorNumber As Long
    Dim stateErrorDescription As String

    If formulaShape Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7493, "VisualTeX", _
            "The selected image conversion target is invalid."
    End If

    ' Direct conversion must not run the normal image-size synchronizer. Word
    ' can invalidate selection-backed collection wrappers while a Ribbon
    ' callback is active, which previously raised Error 5941 before the hidden
    ' imageToNative Session was written. Read the persisted scale state only;
    ' when an older image has no state, derive it from its stable geometry.
    On Error Resume Next
    Err.Clear
    stateReadable = VTTryReadWordImageScaleState( _
        formulaShape.Range.Document, formulaId, fontSizePt, _
        referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
        observedWordFontSizePt)
    Err.Clear
    On Error GoTo StateFailed

    If stateReadable And _
       VTValidWordFormulaFontSize(fontSizePt) And _
       referenceWidthPt > 0# And referenceHeightPt > 0# Then
        If Not VTValidWordFormulaFontSize(observedWordFontSizePt) Then
            observedWordFontSizePt = fontSizePt
        End If
        Exit Sub
    End If

    fontSizePt = VTInlineShapeWordFontSize(formulaShape)
    If Not VTValidWordFormulaFontSize(fontSizePt) Then
        fontSizePt = VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
    End If
    observedWordFontSizePt = fontSizePt
    referenceWidthPt = formulaShape.Width * _
        VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT / fontSizePt
    referenceHeightPt = VTWordImageGlyphHeight(formulaShape) * _
        VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT / fontSizePt
    referenceBaselinePt = VTInlineShapeFontPosition(formulaShape) * _
        VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT / fontSizePt
    If referenceBaselinePt < -256# Then referenceBaselinePt = -256#
    If referenceBaselinePt > 0# Then referenceBaselinePt = 0#
    If referenceWidthPt <= 0# Or referenceHeightPt <= 0# Then
        Err.Raise vbObjectError + 7493, "VisualTeX", _
            "The selected image conversion geometry is invalid."
    End If
    Exit Sub

StateFailed:
    stateErrorNumber = Err.Number
    stateErrorDescription = Err.Description
    If stateErrorNumber = 0 Then stateErrorNumber = vbObjectError + 7493
    Err.Raise stateErrorNumber, "VisualTeX image conversion state", _
        stateErrorDescription
End Sub

Public Sub VTPrepareLegacyNativeMathForEdit(ByRef nativeMath As OMath)
    Dim legacy As Bookmark
    Dim id As String
    Dim metadata As String
    Dim live As Range
    If nativeMath Is Nothing Then Exit Sub
    Set live = nativeMath.Range.Duplicate
    If Not VTTryFindNativeFormulaBookmarkLocally(live, legacy) Then Exit Sub
    id = VTFormulaIdFromBookmarkSuffix(Mid$(legacy.Name, Len(VT_WORD_NATIVE_BOOKMARK_PREFIX) + 1))
    If Not VTTryReadWordMetadataPayload(live.Document, id, metadata) Then Exit Sub
    If Not VTIsEncodedMetadata(metadata) Then Exit Sub
    ' Older managed native inserts imposed Caption on numbered math. Retire
    ' that known plugin-owned style assignment; explicit per-run colours stay.
    If Not VTPureNativeNumberField(live) Is Nothing Then
        If StrComp(live.Paragraphs(1).Range.Style.NameLocal, _
           live.Document.Styles(wdStyleCaption).NameLocal, vbTextCompare) = 0 Then _
            live.Paragraphs(1).Range.Style = wdStyleNormal
    End If
    VTFinishPureNativeEquation live, id
    Set nativeMath = live.OMaths(1)
End Sub

' Unique public entry points used by real-Word acceptance runs. The production
' Startup add-in can remain loaded while an isolated candidate is exercised;
' Word otherwise hides duplicate macro names from the Macros dialog.
Public Function VTWordConvertNativeBookmarkToImageFast( _
    ByVal nativeBookmark As Bookmark, _
    Optional ByRef performancePhases As Variant) As Boolean

    Dim targetDocument As Document
    Dim nativeMath As OMath
    Dim sourceRange As Range
    Dim insertionRange As Range
    Dim candidate As InlineShape
    Dim sequenceField As Field
    Dim visibleNumberField As Field
    Dim numberLayoutRange As Range
    Dim numberBookmarks As Collection
    Dim undoRecord As Object
    Dim formulaId As String
    Dim nativeBookmarkName As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim formulaReference As String
    Dim latexBase64 As String
    Dim captionText As String
    Dim preservedNumberText As String
    Dim vectorDocumentPath As String
    Dim fallbackImagePath As String
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim observedWordFontSizePt As Double
    Dim widthPoints As Double
    Dim heightPoints As Double
    Dim baselinePoints As Double
    Dim cachedWidthPoints As Double
    Dim cachedHeightPoints As Double
    Dim cachedBaselinePoints As Double
    Dim useCachedNaturalGeometry As Boolean
    Dim internalMutationStarted As Boolean
    Dim undoRecordStarted As Boolean
    Dim sourceDeleted As Boolean
    Dim conversionErrorNumber As Long
    Dim conversionErrorDescription As String
    Dim phaseStartedAt As Single
    Dim resolveSeconds As Double
    Dim metadataSeconds As Double
    Dim cacheSeconds As Double
    Dim transactionSeconds As Double
    Dim stagingSeconds As Double
    Dim finalizeSeconds As Double

    phaseStartedAt = Timer
    If nativeBookmark Is Nothing Then Exit Function
    nativeBookmarkName = nativeBookmark.Name
    If Not VTTryFormulaIdFromNativeBookmark( _
       nativeBookmarkName, formulaId) Then Exit Function
    Set targetDocument = nativeBookmark.Range.Document
    Set nativeMath = VTNativeMathForBookmark(nativeBookmark)
    If nativeMath Is Nothing Then Exit Function
    If Not VTTryReadWordFormulaFormat( _
       targetDocument, formulaId, displayMode, numbered) Then Exit Function
    If numbered Then
        ' Only a healthy canonical Equation-array formula is eligible for the
        ' local conversion. Legacy table-numbered OMML stays on the compatibility
        ' path so it can be migrated without weakening the fast-path invariants.
        If displayMode <> "block" Or _
           nativeMath.Range.Information(wdWithInTable) Then Exit Function
        Set sequenceField = VTPureNativeNumberField( _
            nativeMath.Range.Duplicate)
        If sequenceField Is Nothing Then Exit Function
        preservedNumberText = VTEquationNumberTextForFormula( _
            targetDocument, formulaId)
        If Len(preservedNumberText) = 0 Then Exit Function
        Set numberBookmarks = VTCaptureNativeNumberBookmarks( _
            nativeMath.Range.Duplicate)
    End If
    resolveSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    phaseStartedAt = Timer
    If Not VTTryReadWordMetadataPayload( _
       targetDocument, formulaId, encodedMetadata) Then Exit Function
    If Not VTWordNativeSignatureMatches( _
       targetDocument, formulaId, nativeMath) Then Exit Function
    metadataSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    phaseStartedAt = Timer
    If Not VTTryReadWordImageScaleState( _
       targetDocument, formulaId, fontSizePt, referenceWidthPt, _
       referenceHeightPt, referenceBaselinePt, observedWordFontSizePt) Then
        Exit Function
    End If
    useCachedNaturalGeometry = _
        VTWordImageScaleStateUsesPlaceholderGeometry( _
            referenceWidthPt, referenceHeightPt, referenceBaselinePt)
    vectorDocumentPath = VTImageCacheDocumentPath(formulaId)
    If Not VTPathFileExists(vectorDocumentPath) Then Exit Function
    fallbackImagePath = VTImageCachePngPath(formulaId)

    widthPoints = referenceWidthPt * fontSizePt / _
        VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
    heightPoints = referenceHeightPt * fontSizePt / _
        VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
    baselinePoints = referenceBaselinePt * fontSizePt / _
        VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
    If widthPoints <= 0# Or heightPoints <= 0# Or _
       baselinePoints < -256# Or baselinePoints > 0# Then Exit Function
    formulaReference = VTFormulaReference( _
        formulaId, displayMode, numbered)
    captionText = "VisualTeX formula"
    If numbered Then
        If VTTryReadWordLatexPayload( _
           targetDocument, formulaId, latexBase64) Then
            captionText = VTEquationCrossReferenceText(latexBase64)
        End If
    End If
    Set sourceRange = nativeMath.Range.Duplicate
    cacheSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    phaseStartedAt = Timer

    On Error GoTo ConversionFailed
    Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
    undoRecord.StartCustomRecord "VisualTeX Convert OMML to Image"
    undoRecordStarted = True
    VTBeginWordInternalMutation
    internalMutationStarted = True
    transactionSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    phaseStartedAt = Timer
    ' Replace the exact OMath range, then run the same detachment transaction as
    ' the full native-to-image path. A collapsed Range at OMath.Range.End stays
    ' inside the math zone on Word for Mac; the previous fast path therefore
    ' inserted the SVG into the old OMath and could leave a visible math frame
    ' around the picture after conversion. Replacing the native range first
    ' avoids manufacturing that hybrid container, while the Undo record remains
    ' the rollback copy for any later validation failure.
    Set insertionRange = sourceRange.Duplicate
    Set candidate = VTAddWordFormulaPicture( _
        targetDocument, insertionRange, vectorDocumentPath, fallbackImagePath)
    sourceDeleted = True
    Set candidate = VTDetachWordFormulaPictureFromMath(candidate)
    cachedWidthPoints = CDbl(candidate.Width)
    cachedHeightPoints = CDbl(candidate.Height)
    cachedBaselinePoints = VTInlineShapeFontPosition(candidate)
    If useCachedNaturalGeometry Then
        ' Document-imported OMML from older releases stored 14 x 14 as a scale
        ' sentinel even though its durable image DOCX had the correct drawing
        ' extent. Adopt that real Word picture geometry before any resize. This
        ' also repairs the scale variable for every later round trip.
        If cachedWidthPoints <= 0# Or cachedHeightPoints <= 0# Or _
           cachedWidthPoints > 10000# Or cachedHeightPoints > 10000# Then
            Err.Raise vbObjectError + 7422, "VisualTeX", _
                "The cached VisualTeX image has invalid natural dimensions."
        End If
        widthPoints = cachedWidthPoints
        heightPoints = cachedHeightPoints
        If displayMode = "inline" Then
            baselinePoints = cachedBaselinePoints
        Else
            baselinePoints = 0#
        End If
        referenceWidthPt = widthPoints * _
            VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT / fontSizePt
        referenceHeightPt = heightPoints * _
            VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT / fontSizePt
        referenceBaselinePt = baselinePoints * _
            VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT / fontSizePt
    End If
    ' The reference geometry is size independent; draw at the global size.
    If Abs(VTConversionFontSize(displayMode, fontSizePt) - fontSizePt) > 0.001 Then
        fontSizePt = VTConversionFontSize(displayMode, fontSizePt)
        widthPoints = referenceWidthPt * fontSizePt / _
            VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
        heightPoints = referenceHeightPt * fontSizePt / _
            VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
        baselinePoints = referenceBaselinePt * fontSizePt / _
            VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
    End If
    stagingSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    phaseStartedAt = Timer
    candidate.LockAspectRatio = msoFalse
    candidate.Width = CSng(widthPoints)
    candidate.Height = CSng(heightPoints)
    candidate.LockAspectRatio = msoTrue
    candidate.AlternativeText = encodedMetadata
    candidate.Title = formulaReference
    On Error Resume Next
    candidate.Range.Font.Size = CSng(fontSizePt)
    Err.Clear
    On Error GoTo ConversionFailed
    If displayMode = "inline" Then
        VTApplyWordInlineImageBaseline _
            candidate, referenceHeightPt, referenceBaselinePt, fontSizePt
    Else
        candidate.Range.Font.Position = 0
    End If
    If Abs(candidate.Width - widthPoints) > 0.1 Or _
       Abs(VTWordImageGlyphHeight(candidate) - heightPoints) > 0.1 Or _
       candidate.AlternativeText <> encodedMetadata Or _
       candidate.Title <> formulaReference Then
        Err.Raise vbObjectError + 7422, "VisualTeX", _
            "Word did not persist the cached VisualTeX image properties."
    End If

    If targetDocument.Bookmarks.Exists(nativeBookmarkName) Then
        targetDocument.Bookmarks(nativeBookmarkName).Delete
    End If
    Set candidate = VTEnsureVisualTeXImageMacroButton(candidate)
    VTDeleteTrailingNativeImageArtifact candidate, preservedNumberText
    If displayMode = "inline" Then
        VTApplyWordInlineImageBaseline _
            candidate, referenceHeightPt, referenceBaselinePt, fontSizePt
    ElseIf numbered Then
        ' The native formula and image formula own independent SEQ systems.
        ' Rebuild this image's private sequence directly; never carry the OMML
        ' sequence into the image paragraph or manufacture a legacy REF helper.
        Set numberLayoutRange = VTWriteSingleParagraphImageNumber( _
            candidate, formulaId)
        Set candidate = VTResolveImageFormulaByIdentity( _
            targetDocument, formulaId, encodedMetadata, _
            formulaReference, numberLayoutRange.Start)
        VTRestoreImageNumberBookmarks candidate, numberBookmarks
        Set visibleNumberField = VTImageSequenceField( _
            VTImagePlaceRef(candidate.Range), False)
        If visibleNumberField Is Nothing Or _
           VTComparableEquationNumberText( _
               VTCurrentImageNumberRange( _
                   VTImagePlaceRef(candidate.Range)).Text) <> _
           VTComparableEquationNumberText(preservedNumberText) Then
            Err.Raise vbObjectError + 7422, "VisualTeX", _
                "The converted image did not preserve its Equation ordinal."
        End If
    Else
        VTNormalizeUnnumberedDisplayParagraph candidate.Range
        VTNormalizeImageDisplayParagraph candidate.Range
    End If
    VTSetWordImageOwnerBookmark targetDocument, candidate, formulaId
    If useCachedNaturalGeometry Then
        VTSetWordImageScaleState _
            targetDocument, formulaId, fontSizePt, referenceWidthPt, _
            referenceHeightPt, referenceBaselinePt, fontSizePt
    End If
    candidate.Select

    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    If undoRecordStarted Then
        undoRecord.EndCustomRecord
        undoRecordStarted = False
    End If
    finalizeSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    If Not IsMissing(performancePhases) Then
        performancePhases = Array( _
            resolveSeconds, metadataSeconds, cacheSeconds, _
            transactionSeconds, stagingSeconds, finalizeSeconds)
    End If
    VTWordConvertNativeBookmarkToImageFast = True
    Exit Function

ConversionFailed:
    conversionErrorNumber = Err.Number
    conversionErrorDescription = Err.Description
    Resume ConversionRollback

ConversionRollback:
    On Error Resume Next
    If internalMutationStarted Then VTEndWordInternalMutation
    If undoRecordStarted Then undoRecord.EndCustomRecord
    If sourceDeleted Then
        targetDocument.Undo 1
    ElseIf Not candidate Is Nothing Then
        candidate.Delete
    End If
    ' Undo can restore only part of Word's field/bookmark graph after a native
    ' carrier replacement fails. Reconcile the formula lifecycle immediately so
    ' an external VT_N_/SEQ helper cannot outlive its visible formula and poison
    ' the next conversion or numbering retry.
    VTPruneOrphanedEquationNumberScaffolds targetDocument
    On Error GoTo 0
    Err.Raise conversionErrorNumber, _
        "VisualTeX Word cached native-to-image conversion", _
        conversionErrorDescription
End Function

Public Function VTMigrateLegacyNumberedImageTableToParagraph( _
    ByRef formulaShape As InlineShape, _
    ByVal formulaId As String, _
    ByVal captionText As String) As Range

    Dim documentObject As Document
    Dim layoutTable As Table
    Dim centerRange As Range
    Dim convertedRange As Range
    Dim paragraphRange As Range
    Dim migratedRange As Range
    Dim sequenceFieldBefore As Field
    Dim sequenceFieldAfter As Field
    Dim tableStart As Long
    Dim sourceTitle As String
    Dim sourceAlternativeText As String
    Dim sourceWidth As Single
    Dim sourceHeight As Single
    Dim sourceLockAspectRatio As Long
    Dim helperWasPresent As Boolean

    If formulaShape Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7569, "VisualTeX", _
            "The legacy numbered image table target is invalid."
    End If
    If Not formulaShape.Range.Information(wdWithInTable) Then
        Set VTMigrateLegacyNumberedImageTableToParagraph = _
            formulaShape.Range.Paragraphs(1).Range.Duplicate
        Exit Function
    End If
    If formulaShape.Range.Tables.Count <> 1 Then
        Err.Raise vbObjectError + 7569, "VisualTeX", _
            "The legacy numbered image belongs to an ambiguous table."
    End If

    Set documentObject = formulaShape.Range.Document
    Set layoutTable = formulaShape.Range.Tables(1)
    If layoutTable.Rows.Count <> 1 Or layoutTable.Columns.Count <> 3 Then
        Err.Raise vbObjectError + 7569, "VisualTeX", _
            "Only the legacy one-row, three-column image Equation table can be migrated."
    End If
    Set centerRange = layoutTable.Cell(1, 2).Range.Duplicate
    If centerRange.InlineShapes.Count <> 1 Or centerRange.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7569, "VisualTeX", _
            "The legacy image Equation table has no unique picture in its center cell."
    End If
    Set formulaShape = centerRange.InlineShapes(1)
    If Not VTIsVisualTeXInlineShape(formulaShape) Then
        Err.Raise vbObjectError + 7569, "VisualTeX", _
            "The legacy numbered table does not contain a VisualTeX image formula."
    End If

    ' Keep the durable VT_N_/VT_C_ helper whenever it already exists. With that
    ' identity preserved VTInsertEquationNumber sees sequenceOrderChanged=False,
    ' so an edit/migration performs only formula-local work instead of replaying
    ' every Equation in the document.
    Set sequenceFieldBefore = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    helperWasPresent = Not sequenceFieldBefore Is Nothing

    tableStart = layoutTable.Range.Start
    sourceTitle = formulaShape.Title
    sourceAlternativeText = formulaShape.AlternativeText
    sourceWidth = formulaShape.Width
    sourceHeight = formulaShape.Height
    sourceLockAspectRatio = formulaShape.LockAspectRatio

    ' Let Word itself flatten the legacy 1x3 structure. For an empty left cell,
    ' image center cell and numbered right cell this produces exactly one plain
    ' paragraph separated by ordinary TAB characters. Fields, pictures and
    ' Bookmarks survive this operation much more reliably than copying cells out
    ' through temporary documents.
    Set convertedRange = layoutTable.ConvertToText( _
        Separator:=wdSeparateByTabs)
    Set paragraphRange = documentObject.Range( _
        Start:=tableStart, End:=tableStart).Paragraphs(1).Range.Duplicate
    If paragraphRange.Information(wdWithInTable) Or _
       paragraphRange.InlineShapes.Count <> 1 Or _
       paragraphRange.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7569, "VisualTeX", _
            "Word did not flatten the legacy numbered image table into one ordinary paragraph."
    End If

    Set formulaShape = paragraphRange.InlineShapes(1)
    formulaShape.LockAspectRatio = msoFalse
    formulaShape.Width = sourceWidth
    formulaShape.Height = sourceHeight
    formulaShape.LockAspectRatio = sourceLockAspectRatio
    formulaShape.Title = sourceTitle
    formulaShape.AlternativeText = sourceAlternativeText
    formulaShape.Range.Font.Position = 0
    Set formulaShape = VTEnsureVisualTeXImageMacroButton(formulaShape)

    ' Reuse the production paragraph normalizer. It removes the tabs/visible REF
    ' created by ConvertToText and rebuilds the single canonical layout:
    ' <TAB><image><TAB>(number)<CR> plus the external native SEQ helper.
    Set migratedRange = VTInsertEquationNumber( _
        formulaShape, formulaId, captionText)
    Set formulaShape = VTResolveImageFormulaByIdentity( _
        documentObject, formulaId, sourceAlternativeText, _
        sourceTitle, migratedRange.Start)
    If migratedRange.Information(wdWithInTable) Or _
       formulaShape.Range.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 7569, "VisualTeX", _
            "The migrated image Equation unexpectedly remained in a table."
    End If

    If helperWasPresent Then
        Set sequenceFieldAfter = VTNativeEquationSequenceHelperField( _
            documentObject, formulaId)
        If sequenceFieldAfter Is Nothing Then
            Err.Raise vbObjectError + 7569, "VisualTeX", _
                "The legacy image migration lost its existing Equation SEQ identity."
        End If
    End If
    Set VTMigrateLegacyNumberedImageTableToParagraph = _
        migratedRange.Duplicate
End Function

Private Function VTMigrateLegacyNumberedImageTables( _
    ByVal documentObject As Document) As Long

    Dim layoutTable As Table
    Dim centerRange As Range
    Dim formulaShape As InlineShape
    Dim migratedRange As Range
    Dim formulaId As String
    Dim tableIndex As Long

    If documentObject Is Nothing Then Exit Function
    For tableIndex = documentObject.Tables.Count To 1 Step -1
        If tableIndex > documentObject.Tables.Count Then GoTo NextTable
        Set layoutTable = documentObject.Tables(tableIndex)
        If layoutTable.Rows.Count <> 1 Or _
           layoutTable.Columns.Count <> 3 Then GoTo NextTable
        Set centerRange = layoutTable.Cell(1, 2).Range.Duplicate
        If centerRange.InlineShapes.Count <> 1 Or _
           centerRange.OMaths.Count <> 0 Then GoTo NextTable
        formulaId = VTFormulaIdForValidNumberedTable(layoutTable)
        If Len(formulaId) = 0 Then GoTo NextTable
        Set formulaShape = centerRange.InlineShapes(1)
        If Not VTIsVisualTeXInlineShape(formulaShape) Then GoTo NextTable
        Set migratedRange = _
            VTMigrateLegacyNumberedImageTableToParagraph( _
                formulaShape, formulaId, "VisualTeX formula")
        VTMigrateLegacyNumberedImageTables = _
            VTMigrateLegacyNumberedImageTables + 1
NextTable:
    Next tableIndex
End Function

Private Function VTNormalizeImageConversionCarrier( _
    ByVal sourceImage As InlineShape, _
    ByVal formulaId As String, _
    ByRef targetDocumentMutated As Boolean) As InlineShape

    Dim containerRange As Range
    Dim candidateField As Field
    Dim remainingText As String

    Set VTNormalizeImageConversionCarrier = sourceImage
    If sourceImage.Range.OMaths.Count = 0 Then Exit Function
    If sourceImage.Range.OMaths.Count <> 1 Then GoTo InvalidCarrier
    Set containerRange = VTNativeMathRoot(sourceImage.Range.OMaths(1)).Range.Duplicate
    If containerRange.InlineShapes.Count <> 1 Then GoTo InvalidCarrier
    ' Older native-to-image replacement left # / (REF) around the image inside
    ' OMath. This is an image carrier, not a native source to build or number.
    ' Accept only this image and its own number; never delete surrounding math.
    remainingText = containerRange.Text
    remainingText = Replace$(remainingText, sourceImage.Range.Text, "", 1, 1)
    For Each candidateField In containerRange.Fields
        If candidateField.Type <> wdFieldRef Then GoTo InvalidCarrier
        If StrComp(VTReferenceTargetBookmarkName(candidateField.Code.Text), _
           VTEquationSequenceNumberBookmarkName(formulaId), vbTextCompare) <> 0 Then _
            GoTo InvalidCarrier
        remainingText = Replace$(remainingText, candidateField.Result.Text, "", 1, 1)
    Next candidateField
    remainingText = Replace$(remainingText, "#", "")
    remainingText = Replace$(remainingText, "(", "")
    remainingText = Replace$(remainingText, ")", "")
    remainingText = Replace$(remainingText, " ", "")
    remainingText = Replace$(remainingText, vbTab, "")
    remainingText = Replace$(remainingText, vbCr, "")
    remainingText = Replace$(remainingText, ChrW(8203), "")
    remainingText = Replace$(remainingText, ChrW(8288), "")
    If Len(remainingText) <> 0 Then GoTo InvalidCarrier
    Set sourceImage = VTDetachWordFormulaPictureFromMath( _
        sourceImage, targetDocumentMutated)
    VTSetWordImageOwnerBookmark sourceImage.Range.Document, sourceImage, formulaId
    Set VTNormalizeImageConversionCarrier = sourceImage
    Exit Function
InvalidCarrier:
    Err.Raise vbObjectError + 7564, "VisualTeX", _
        "The image conversion carrier contains unrelated mathematics or fields."
End Function

Public Function VTPrepareNumberedImageCarrierForNativeFastConversion( _
    ByVal documentObject As Document, _
    ByVal formulaShape As InlineShape, _
    ByVal formulaId As String) As InlineShape

    Dim paragraphRange As Range
    Dim formulaContainerRange As Range
    Dim prefixRange As Range
    Dim suffixRange As Range
    Dim numberRange As Range
    Dim visibleNumberField As Field
    Dim candidateField As Field
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim formulaStart As Long
    Dim imageMetadata As String
    Dim imageFormulaReference As String
    Dim tailText As String

    If documentObject Is Nothing Or formulaShape Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7430, "VisualTeX", _
            "The numbered image conversion carrier is invalid."
    End If
    If formulaShape.Range.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 7430, "VisualTeX", _
            "The numbered image conversion carrier cannot be inside a table."
    End If
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    If paragraphRange.InlineShapes.Count <> 1 Or _
       paragraphRange.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7430, "VisualTeX", _
            "The numbered image conversion carrier is ambiguous."
    End If

    sequenceBookmarkName = VTEquationSequenceNumberBookmarkName(formulaId)
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    Set visibleNumberField = VTImageEquationReferenceField( _
        formulaShape.Range.Duplicate, formulaId)
    If visibleNumberField Is Nothing Or _
       Not documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        Err.Raise vbObjectError + 7430, "VisualTeX", _
            "The numbered image conversion carrier has no owned Equation REF."
    End If
    Set numberRange = VTImageEquationNumberRange( _
        formulaShape.Range.Duplicate, visibleNumberField)
    If numberRange Is Nothing Then
        Err.Raise vbObjectError + 7430, "VisualTeX", _
            "The numbered image conversion carrier has no complete visible number."
    End If

    imageMetadata = formulaShape.AlternativeText
    imageFormulaReference = formulaShape.Title
    Set formulaContainerRange = VTVisualTeXImageContainerRange(formulaShape)
    formulaStart = formulaContainerRange.Start
    Set prefixRange = documentObject.Range( _
        Start:=paragraphRange.Start, End:=formulaContainerRange.Start)
    If prefixRange.Fields.Count <> 0 Or _
       prefixRange.InlineShapes.Count <> 0 Or _
       prefixRange.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(prefixRange) Then
        Err.Raise vbObjectError + 7430, "VisualTeX", _
            "The numbered image has body content before its formula."
    End If

    Set suffixRange = documentObject.Range( _
        Start:=formulaContainerRange.End, End:=paragraphRange.End - 1)
    If VTIsImagePlaceRef(visibleNumberField) Then
        If VTEquationFieldStart(visibleNumberField) < suffixRange.Start Or _
           VTEquationFieldEnd(visibleNumberField) > suffixRange.End Then
            Err.Raise vbObjectError + 7430, "VisualTeX", "The image number crosses its paragraph: field=" & CStr(VTEquationFieldStart(visibleNumberField)) & "-" & CStr(VTEquationFieldEnd(visibleNumberField)) & "; tail=" & CStr(suffixRange.Start) & "-" & CStr(suffixRange.End)
        End If
    Else
    For Each candidateField In suffixRange.Fields
        If candidateField.Type <> wdFieldRef Or _
           StrComp(VTReferenceTargetBookmarkName( _
               candidateField.Code.Text), sequenceBookmarkName, _
               vbTextCompare) <> 0 Or _
           VTEquationFieldStart(candidateField) <> _
               VTEquationFieldStart(visibleNumberField) Then
            Err.Raise vbObjectError + 7430, "VisualTeX", _
                "The numbered image has an unrelated trailing Word field."
        End If
    Next candidateField
    If suffixRange.Fields.Count <> 1 Then
        Err.Raise vbObjectError + 7430, "VisualTeX", _
            "The numbered image does not have exactly one visible Equation REF."
    End If
    End If
    tailText = documentObject.Range(suffixRange.Start, _
        VTEquationFieldStart(visibleNumberField)).Text & _
        documentObject.Range(VTEquationFieldEnd(visibleNumberField), suffixRange.End).Text
    tailText = Replace$(tailText, vbTab, "")
    tailText = Replace$(tailText, "(", "")
    tailText = Replace$(tailText, ")", "")
    tailText = Replace$(tailText, " ", "")
    tailText = Replace$(tailText, ChrW(160), "")
    tailText = Replace$(tailText, ChrW(8203), "")
    tailText = Replace$(tailText, ChrW(8288), "")
    If Len(tailText) <> 0 Then
        Err.Raise vbObjectError + 7430, "VisualTeX", _
            "The numbered image has body content after its formula."
    End If

    ' Keep VT_N_/VT_C_ and their external SEQ untouched. Only the carrier's
    ' visible VT_R_ row is removed so InsertFile can target a field-free
    ' paragraph. The caller's UndoRecord owns rollback from this first write.
    If documentObject.Bookmarks.Exists(numberBookmarkName) Then
        documentObject.Bookmarks(numberBookmarkName).Delete
    End If
    visibleNumberField.Unlink
    Set formulaShape = VTResolveImageFormulaByIdentity( _
        documentObject, formulaId, imageMetadata, _
        imageFormulaReference, formulaStart)
    Set formulaContainerRange = VTVisualTeXImageContainerRange(formulaShape)
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    Set suffixRange = documentObject.Range( _
        Start:=formulaContainerRange.End, End:=paragraphRange.End - 1)
    If suffixRange.End > suffixRange.Start Then suffixRange.Delete
    Set formulaShape = VTResolveImageFormulaByIdentity( _
        documentObject, formulaId, imageMetadata, _
        imageFormulaReference, formulaStart)
    Set formulaContainerRange = VTVisualTeXImageContainerRange(formulaShape)
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    Set prefixRange = documentObject.Range( _
        Start:=paragraphRange.Start, End:=formulaContainerRange.Start)
    If prefixRange.End > prefixRange.Start Then prefixRange.Delete

    Set formulaShape = VTResolveImageFormulaByIdentity( _
        documentObject, formulaId, imageMetadata, _
        imageFormulaReference, paragraphRange.Start)
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    If paragraphRange.Fields.Count <> 0 Or _
       paragraphRange.OMaths.Count <> 0 Or _
       paragraphRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7430, "VisualTeX", _
            "Word did not prepare a clean numbered image conversion carrier" & _
            " [fields=" & CStr(paragraphRange.Fields.Count) & _
            "; omaths=" & CStr(paragraphRange.OMaths.Count) & _
            "; images=" & CStr(paragraphRange.InlineShapes.Count) & "]."
    End If
    Set VTPrepareNumberedImageCarrierForNativeFastConversion = formulaShape
End Function

Public Function VTWordConvertInlineShapeToNativeFast( _
    ByVal target As InlineShape, _
    Optional ByRef conversionPath As String, _
    Optional ByRef performancePhases As Variant) As Boolean

    Dim numberBookmarks As Collection
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim ommlBase64 As String
    Dim latexBase64 As String
    Dim captionText As String
    Dim preservedNumberText As String
    Dim nativeDocumentPath As String
    Dim targetDocument As Document
    Dim sourceContainerRange As Range
    Dim insertionRange As Range
    Dim equationRange As Range
    Dim numberLayoutRange As Range
    Dim nativeMath As OMath
    Dim sourceImage As InlineShape
    Dim sequenceField As Field
    Dim visibleNumberField As Field
    Dim undoRecord As Object
    Dim sourceFontSizePt As Double
    Dim sourceHeightPoints As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim observedWordFontSizePt As Double
    Dim sourceStart As Long
    Dim equationStart As Long
    Dim internalMutationStarted As Boolean
    Dim undoRecordStarted As Boolean
    Dim mutationOccurred As Boolean
    Dim numberCreated As Boolean
    Dim directInsertPhases As Variant
    Dim directInsertDiagnostic As String
    Dim conversionErrorNumber As Long
    Dim conversionErrorDescription As String
    Dim phaseStartedAt As Single
    Dim prepareSeconds As Double
    Dim insertSeconds As Double
    Dim deleteSeconds As Double
    Dim numberSeconds As Double
    Dim finalizeSeconds As Double

    phaseStartedAt = Timer
    conversionPath = "unavailable"
    If target Is Nothing Or Not VTIsVisualTeXInlineShape(target) Then Exit Function
    If Not VTTryParseFormulaReference( _
       target.Title, formulaId, displayMode, numbered) Then Exit Function
    Set targetDocument = target.Range.Document
    If numbered Then Set numberBookmarks = VTCaptureImageNumberBookmarks(target, formulaId)
    Set sourceContainerRange = VTVisualTeXImageContainerRange(target)
    ' The slow path repairs a historical image-in-OMath carrier before reading
    ' numbering topology. Never insert cached math into that existing math zone.
    If sourceContainerRange.OMaths.Count <> 0 Then Exit Function
    If numbered And sourceContainerRange.Information(wdWithInTable) Then Exit Function
    If numbered Then
        preservedNumberText = VTEquationNumberTextForFormula(targetDocument, formulaId)
    End If
    If Not VTTryReadWordMetadataPayload( _
       targetDocument, formulaId, encodedMetadata) Then
        encodedMetadata = target.AlternativeText
    End If
    If Not VTIsEncodedMetadata(encodedMetadata) Then Exit Function
    If Not VTTryReadWordOmmlPayload( _
       targetDocument, formulaId, ommlBase64) Then Exit Function
    nativeDocumentPath = VTNativeWordDocumentPath(formulaId)
    If Not VTPathFileExists(nativeDocumentPath) Then
        conversionPath = "cache-unavailable"
        Exit Function
    End If

    VTReadWordImageConversionState _
        target, formulaId, sourceFontSizePt, referenceWidthPt, _
        referenceHeightPt, referenceBaselinePt, observedWordFontSizePt
    sourceFontSizePt = VTConversionFontSize(displayMode, sourceFontSizePt)
    sourceHeightPoints = VTWordImageGlyphHeight(target)
    captionText = "VisualTeX formula"
    If numbered Then
        If VTTryReadWordLatexPayload( _
           targetDocument, formulaId, latexBase64) Then
            captionText = VTEquationCrossReferenceText(latexBase64)
        End If
    End If
    sourceStart = sourceContainerRange.Start
    prepareSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)

    On Error GoTo ConversionFailed
    Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
    undoRecord.StartCustomRecord "VisualTeX Convert Image to OMML"
    undoRecordStarted = True
    VTBeginWordInternalMutation
    internalMutationStarted = True

    phaseStartedAt = Timer
    If numbered Then
        mutationOccurred = True
        Set target = VTPrepareNumberedImageCarrierForNativeFastConversion( _
            targetDocument, target, formulaId)
        Set sourceContainerRange = VTVisualTeXImageContainerRange(target)
        sourceStart = sourceContainerRange.Start
    End If

    ' Insert before the still-live source image. The image remains the rollback
    ' copy until Word has parsed and validated exactly one cached OMath.
    Set insertionRange = targetDocument.Range( _
        Start:=sourceStart, End:=sourceStart)
    ' A display image now occupies a dedicated field-free paragraph. For a
    ' numbered image the local preparation above removed only VT_R_; VT_N_,
    ' VT_C_ and the external SEQ remain live. Import the cached DOCX directly so
    ' neither direction opens a hidden Word document for a format-only change.
    If displayMode = "block" Then
        Set equationRange = VTTryInsertNativeEquationFileFast( _
            sourceContainerRange, nativeDocumentPath, "block", True, _
            directInsertPhases, directInsertDiagnostic)
        If Not equationRange Is Nothing Then
            If numbered Then
                conversionPath = "numbered-direct-cache-insert"
            Else
                conversionPath = "direct-cache-insert"
            End If
        End If
    End If
    If equationRange Is Nothing Then
        conversionPath = "staging-document"
        If Len(directInsertDiagnostic) > 0 Then _
            conversionPath = conversionPath & ":" & directInsertDiagnostic
        Set equationRange = VTInsertCachedNativeEquationFast( _
            insertionRange, nativeDocumentPath, sourceFontSizePt)
    End If
    If equationRange Is Nothing Then GoTo CacheUnavailable
    equationStart = equationRange.Start
    insertSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    phaseStartedAt = Timer

    Set sourceImage = VTWordImageOwnerShape(targetDocument, formulaId)
    If sourceImage Is Nothing Then
        Set sourceImage = VTFindCommittedInlineShapeNearPosition( _
            targetDocument, sourceStart, encodedMetadata, _
            VTFormulaReference(formulaId, displayMode, numbered), 512)
    End If
    If sourceImage Is Nothing Then
        Err.Raise vbObjectError + 7430, "VisualTeX", _
            "Word could not re-resolve the source image after cached OMML insertion."
    End If
    VTDeleteVisualTeXImageContainer sourceImage
    mutationOccurred = True
    deleteSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    phaseStartedAt = Timer

    Set equationRange = VTResolveNativeEquationRange( _
        targetDocument, equationStart, 16)
    If displayMode = "block" Then
        If numbered Then
            numberCreated = False
            Set numberLayoutRange = VTEnsureNativeEquationNumber( _
                equationRange, sourceHeightPoints, formulaId, captionText, numberCreated)
            Set equationRange = VTNumberedFormulaRangeForId( _
                targetDocument, formulaId)
            If equationRange Is Nothing Or _
               equationRange.Information(wdWithInTable) Then
                Err.Raise vbObjectError + 7430, "VisualTeX", _
                    "Word did not create the canonical numbered OMML layout."
            End If
        Else
            Set equationRange = VTPromoteNativeEquationToDisplay(equationRange)
        End If
    Else
        Set equationRange = VTFinalizeInlineNativeEquation(equationRange)
    End If
    numberSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    phaseStartedAt = Timer
    VTV2ApplyEquationTypography equationRange, sourceFontSizePt
    VTSetWordMetadataPayload targetDocument, formulaId, encodedMetadata
    VTSetWordFormulaFormat targetDocument, formulaId, displayMode, numbered
    VTSetWordImageScaleState _
        targetDocument, formulaId, sourceFontSizePt, referenceWidthPt, _
        referenceHeightPt, referenceBaselinePt, sourceFontSizePt
    VTSetNativeFormulaBookmark targetDocument, equationRange, formulaId
    VTDeleteWordImageOwnerBookmark targetDocument, formulaId
    Set nativeMath = VTNativeMathForBookmark( _
        targetDocument.Bookmarks(VTNativeFormulaBookmarkName(formulaId)))
    If nativeMath Is Nothing Then
        Err.Raise vbObjectError + 7430, "VisualTeX", _
            "Word lost the native equation before structure caching."
    End If
    VTRestoreNativeNumberBookmarks equationRange, numberBookmarks
    VTFinishPureNativeEquation equationRange, formulaId
    Set nativeMath = VTNativeMathRoot(equationRange.OMaths(1))
    If nativeMath Is Nothing Then
        Err.Raise vbObjectError + 7430, "VisualTeX", _
            "Word lost the native equation after number finalization."
    End If
    ' Equation format v2: the converted number takes the shared format.
    If numbered And displayMode = "block" Then
        VTV2RestyleNumbers targetDocument
        Set equationRange = VTResolveNativeEquationRange( _
            targetDocument, equationRange.Start, 16)
    End If
    equationRange.Select

    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    If undoRecordStarted Then
        undoRecord.EndCustomRecord
        undoRecordStarted = False
    End If
    finalizeSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    If Not IsMissing(performancePhases) Then
        performancePhases = Array( _
            prepareSeconds, insertSeconds, deleteSeconds, _
            numberSeconds, finalizeSeconds)
    End If
    VTWordConvertInlineShapeToNativeFast = True
    Exit Function

CacheUnavailable:
    conversionPath = "cache-unavailable"
    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    If undoRecordStarted Then
        undoRecord.EndCustomRecord
        undoRecordStarted = False
    End If
    If mutationOccurred Then targetDocument.Undo 1
    Exit Function

ConversionFailed:
    conversionErrorNumber = Err.Number
    conversionErrorDescription = Err.Description
    Resume ConversionRollback

ConversionRollback:
    On Error Resume Next
    If internalMutationStarted Then VTEndWordInternalMutation
    If undoRecordStarted Then undoRecord.EndCustomRecord
    If mutationOccurred Then
        targetDocument.Undo 1
    ElseIf Not equationRange Is Nothing Then
        equationRange.Delete
    End If
    VTPruneOrphanedEquationNumberScaffolds targetDocument
    On Error GoTo 0
    Err.Raise conversionErrorNumber, _
        "VisualTeX Word cached image-to-native conversion", _
        conversionErrorDescription
End Function

Public Sub VTWordConvertInlineShapeToNativeEquation( _
    ByVal target As InlineShape, _
    Optional ByVal regressionFailureAfterSourceRemoval As Boolean = False)
    Dim numberBookmarks As Collection
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
    Dim sourceContainerRange As Range
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
    Dim replaceSourceContainer As Boolean
    Dim sourceFontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim observedWordFontSizePt As Double
    Dim convertedSequenceField As Field
    Dim sourceParagraph As Range
    Dim sourcePrefix As Range
    Dim sourceSuffix As Range
    Dim ownedNumberTail As Range
    Dim undoRecord As Object
    Dim undoRecordStarted As Boolean
    Dim targetDocumentMutated As Boolean
    Dim nativeEquationInserted As Boolean
    Dim conversionStage As String

    conversionStage = "validate-source-image"
    On Error GoTo ConversionFailed
    If target Is Nothing Or Not VTIsVisualTeXInlineShape(target) Then
        Err.Raise vbObjectError + 7430, "VisualTeX", "The selected object is not a VisualTeX formula image."
    End If
    If Not VTTryParseFormulaReference(target.Title, formulaId, displayMode, numbered) Then
        Err.Raise vbObjectError + 7431, "VisualTeX", "The selected VisualTeX formula reference is invalid."
    End If
    Set targetDocument = target.Range.Document
    If numbered Then Set numberBookmarks = VTCaptureImageNumberBookmarks(target, formulaId)
    If Not VTTryReadWordOmmlPayload(targetDocument, formulaId, ommlBase64) Then
        Err.Raise vbObjectError + 7432, "VisualTeX", _
            "This formula has no structural OMML payload. Edit and save it once in the current VisualTeX, then convert it again."
    End If
    nativeDocumentPath = VTNativeWordDocumentPath(formulaId)
    If Not VTPathFileExists(nativeDocumentPath) Then
        ' Older image formulas can retain valid LaTeX/OMML metadata while their
        ' durable native staging DOCX is absent. Rebuild that DOCX through the
        ' resident renderer as a silent image-to-native operation; the frontend
        ' commits automatically and this same Word callback performs the final
        ' replacement without ever revealing the formula editor.
        VTWordEditInlineShape target, True, "imageToNative"
        Exit Sub
    End If

    conversionStage = "begin-conversion-transaction"
    Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
    undoRecord.StartCustomRecord "VisualTeX Convert Image to OMML"
    undoRecordStarted = True
    VTBeginWordInternalMutation
    internalMutationStarted = True
    conversionStage = "normalize-image-carrier"
    Set target = VTNormalizeImageConversionCarrier( _
        target, formulaId, targetDocumentMutated)
    conversionStage = "inspect-numbered-source"
    If numbered And displayMode = "block" And _
       Not target.Range.Information(wdWithInTable) Then
        ' Reject real body text before any image refresh/reconciliation can
        ' replace the visible number. Legacy direct-SEQ tails are migrated by
        ' the existing conversion transaction below.
        Set sourceContainerRange = VTVisualTeXImageContainerRange(target)
        Set sourceParagraph = sourceContainerRange.Paragraphs(1).Range.Duplicate
        Set sourcePrefix = sourceParagraph.Duplicate
        sourcePrefix.End = sourceContainerRange.Start
        Set sourceSuffix = sourceParagraph.Duplicate
        sourceSuffix.Start = sourceContainerRange.End
        Set ownedNumberTail = VTNativeEquationOwnedImageNumberTail( _
            sourceContainerRange, formulaId)
        If Not ownedNumberTail Is Nothing Then sourceSuffix.Start = ownedNumberTail.End
        Set convertedSequenceField = VTFindEquationSequenceField(sourceParagraph)
        If VTWordRangeHasMeaningfulText(sourcePrefix) Or _
           (convertedSequenceField Is Nothing And _
            VTWordRangeHasMeaningfulText(sourceSuffix)) Then
            Err.Raise vbObjectError + 7564, "VisualTeX", _
                "The numbered image has surrounding body text."
        End If
    End If
    If numbered And displayMode = "block" Then
        conversionStage = "resolve-number-sequence"
        Set convertedSequenceField = VTNativeEquationSequenceHelperField( _
            targetDocument, formulaId)
        If Not convertedSequenceField Is Nothing Then
            If InStr(1, convertedSequenceField.Code.Text, _
               VTNativeEquationLabelName(), vbTextCompare) = 0 Then
                ' Reconcile a reopened document's localized labels in document
                ' order before taking source offsets or starting replacement.
                ' Updating just this SEQ first would restart its renamed sequence.
                targetDocumentMutated = True
                VTReconcileEquationNumbers targetDocument
            End If
        End If
    End If
    targetDocumentMutated = True
    If numbered And Not VTImagePlaceRef(target.Range) Is Nothing Then
        Set target = VTPrepareNumberedImageCarrierForNativeFastConversion(targetDocument, target, formulaId)
    End If
    conversionStage = "synchronize-image-state"
    Set target = VTEnsureVisualTeXImageMacroButton(target)
    VTSynchronizeWordImageFormulaShape target
    VTEnsureWordImageScaleState _
        target, formulaId, displayMode, numbered, sourceFontSizePt, _
        referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
        observedWordFontSizePt
    sourceFontSizePt = VTConversionFontSize(displayMode, sourceFontSizePt)
    encodedMetadata = target.AlternativeText
    If VTTryReadWordLatexPayload(targetDocument, formulaId, latexBase64) Then
        captionText = VTEquationCrossReferenceText(latexBase64)
    Else
        captionText = "VisualTeX formula"
    End If
    sourceHeightPoints = VTWordImageGlyphHeight(target)
    nativeDisplayMode = displayMode
    If numbered Or displayMode = "block" Then nativeDisplayMode = "inline"
    Set sourceContainerRange = VTVisualTeXImageContainerRange(target)
    sourceStart = sourceContainerRange.Start
    replaceSourceContainer = (sourceContainerRange.OMaths.Count = 1)
    ' Keep an exact FormattedText backup because display promotion must happen
    ' after the source image is removed. Word shifts adjacent OMath Ranges when
    ' an InlineShape is deleted, so the source cannot safely be deleted without
    ' a rollback copy and a fresh equation lookup.
    conversionStage = "backup-source-image"
    Set sourceBackupDocument = Documents.Add(Visible:=False)
    Set sourceBackupRange = sourceBackupDocument.Content
    sourceBackupRange.Collapse wdCollapseStart
    sourceBackupRange.FormattedText = sourceContainerRange.FormattedText
    If sourceBackupDocument.InlineShapes.Count <> 1 Or _
       Not VTIsVisualTeXInlineShape( _
           sourceBackupDocument.InlineShapes(1)) Then
        Err.Raise vbObjectError + 7472, "VisualTeX", _
            "Word could not back up the VisualTeX image before native conversion."
    End If
    Set sourceBackupRange = _
        VTVisualTeXImageContainerRange( _
            sourceBackupDocument.InlineShapes(1))
    targetDocument.Activate

    If numbered And displayMode = "block" Then
        conversionStage = "resolve-image-after-number-mutation"
        Set sourceImage = VTResolveImageFormulaByIdentity( _
            targetDocument, formulaId, encodedMetadata, _
            VTFormulaReference(formulaId, displayMode, numbered), sourceStart)
        Set convertedSequenceField = VTFindEquationSequenceField( _
            sourceImage.Range.Paragraphs(1).Range.Duplicate)
        If Not convertedSequenceField Is Nothing Then
            ' Schema-6 compatibility input: move the same live SEQ into the
            ' final external helper before native conversion. Healthy current
            ' images already own that helper and skip this migration entirely.
            Set convertedSequenceField = _
                VTMoveDirectImageEquationSequenceToExternalHelper( _
                    sourceImage, formulaId)
        Else
            Set convertedSequenceField = VTNativeEquationSequenceHelperField( _
                targetDocument, formulaId)
            ' If the helper is absent, create it only after replacing the image
            ' with OMath. Do not run the image paragraph numberer as a staging
            ' operation while the source carrier is being converted.
        End If
    End If
    conversionStage = "resolve-image-before-native-insertion"
    Set sourceImage = VTResolveImageFormulaByIdentity( _
        targetDocument, formulaId, encodedMetadata, _
        VTFormulaReference(formulaId, displayMode, numbered), sourceStart)
    Set sourceContainerRange = VTVisualTeXImageContainerRange(sourceImage)
    sourceStart = sourceContainerRange.Start
    replaceSourceContainer = (sourceContainerRange.OMaths.Count = 1)
    Set insertionAnchor = sourceContainerRange.Duplicate
    insertionAnchor.Collapse wdCollapseEnd

    VTSetWordMetadataPayload targetDocument, formulaId, encodedMetadata
    VTSetWordFormulaFormat targetDocument, formulaId, displayMode, numbered
    conversionStage = "insert-native-equation"
    If replaceSourceContainer Then
        Set equationRange = VTInsertNativeEquationAtRange( _
            sourceContainerRange, _
            ommlBase64, _
            nativeDocumentPath, _
            nativeDisplayMode, _
            displayMode = "block", _
            True)
        sourceDeleted = True
    Else
        Set equationRange = VTInsertNativeEquationAtRange( _
            insertionAnchor, _
            ommlBase64, _
            nativeDocumentPath, _
            nativeDisplayMode, _
            displayMode = "block", _
            False)

        conversionStage = "delete-source-image"
        targetDocument.Activate
        Set sourceImage = VTFindUniqueInlineShape(encodedMetadata)
        VTDeleteVisualTeXImageContainer sourceImage
        sourceDeleted = True
    End If
    nativeEquationStart = equationRange.Start
    nativeEquationInserted = True
    If regressionFailureAfterSourceRemoval Then
        conversionStage = "injected-post-delete-failure"
        Err.Raise vbObjectError + 7564, "VisualTeX regression", _
            "Injected image-to-native layout failure after source deletion."
    End If

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
                VTNativeMathRoot(numberLayoutRange.OMaths(1)).Range.Duplicate
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

    On Error Resume Next
    ' This is a conversion of an existing formula, not a new insertion. Select
    ' the converted OMath instead of creating a U+2060 caret anchor after it.
    equationRange.Select
    On Error GoTo ConversionFailed
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
    VTV2ApplyEquationTypography finalFormulaRange, sourceFontSizePt
    VTSetWordImageScaleState _
        targetDocument, formulaId, sourceFontSizePt, referenceWidthPt, _
        referenceHeightPt, referenceBaselinePt, sourceFontSizePt
    VTSetNativeFormulaBookmark _
        targetDocument, finalFormulaRange, formulaId
    VTDeleteWordImageOwnerBookmark targetDocument, formulaId
    If Not targetDocument.Bookmarks.Exists( _
       VTNativeFormulaBookmarkName(formulaId)) Or _
       (numbered And displayMode = "block" And _
        Not targetDocument.Bookmarks.Exists( _
            VTEquationNumberBookmarkName(formulaId))) Then
        Err.Raise vbObjectError + 7460, "VisualTeX", _
            "Word did not preserve the final VisualTeX formula identity."
    End If
    VTRestoreNativeNumberBookmarks finalFormulaRange, numberBookmarks
    VTFinishPureNativeEquation finalFormulaRange, formulaId
    Set finalFormulaRange = VTNativeMathRootRange(finalFormulaRange)
    Set equationRange = finalFormulaRange.Duplicate
    ' Equation format v2: the converted number takes the shared format.
    If numbered And displayMode = "block" Then
        conversionStage = "v2-restyle-number"
        VTV2RestyleNumbers targetDocument
    End If
    sourceBackupDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set sourceBackupDocument = Nothing
    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    If undoRecordStarted Then
        undoRecord.EndCustomRecord
        undoRecordStarted = False
    End If
    VTEnsureOrphanWatchScheduled
    Exit Sub

ConversionFailed:
    conversionErrorNumber = Err.Number
    conversionErrorDescription = conversionStage & ": " & Err.Description
    Resume ConversionRollback

ConversionRollback:
    On Error Resume Next
    ' Roll back the whole conversion, including carrier repair and numbering.
    ' Re-running the image numberer here can fail on the original hybrid math
    ' carrier and cannot restore its original paragraph/bookmark boundaries.
    If internalMutationStarted Then
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    If undoRecordStarted Then
        targetDocument.Activate
        undoRecord.EndCustomRecord
        If targetDocumentMutated Then _
            targetDocument.Undo 1
    End If
    ' Word for Mac can record the source-picture deletion separately from the
    ' preceding OMath insertion even inside one custom Undo record. In that
    ' case one Document.Undo restores the picture but leaves the inserted OMath
    ' beside it. Re-resolve that exact conversion result from its post-insert
    ' anchor and remove only a pure native equation; a restored legacy
    ' image-in-OMath carrier is deliberately preserved.
    If nativeEquationInserted Then
        Set equationRange = Nothing
        Set equationRange = VTResolveNativeEquationRange( _
            targetDocument, nativeEquationStart, 64)
        If Not equationRange Is Nothing Then
            If equationRange.InlineShapes.Count = 0 Then equationRange.Delete
        End If
    End If
    ' Undo may restore the picture without restoring VT_I_. Reacquire it from
    ' immutable metadata and rebuild the owner identity before orphan cleanup.
    Set sourceImage = Nothing
    Set sourceImage = VTWordImageOwnerShape(targetDocument, formulaId)
    If sourceImage Is Nothing Then
        Set sourceImage = VTFindCommittedInlineShapeNearPosition( _
            targetDocument, sourceStart, encodedMetadata, _
            VTFormulaReference(formulaId, displayMode, numbered), 256)
    End If
    If sourceImage Is Nothing And Not sourceBackupRange Is Nothing Then
        Set sourceRestoreRange = targetDocument.Range( _
            Start:=sourceStart, End:=sourceStart)
        sourceRestoreRange.FormattedText = sourceBackupRange.FormattedText
        Set sourceImage = VTFindCommittedInlineShapeNearPosition( _
            targetDocument, sourceStart, encodedMetadata, _
            VTFormulaReference(formulaId, displayMode, numbered), 256)
    End If
    If Not sourceImage Is Nothing Then _
        VTSetWordImageOwnerBookmark targetDocument, sourceImage, formulaId
    If Not sourceBackupDocument Is Nothing Then
        sourceBackupDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    VTPruneOrphanedEquationNumberScaffolds targetDocument
    VTEnsureOrphanWatchScheduled
    On Error GoTo 0
    Err.Raise conversionErrorNumber, "VisualTeX Word image-to-native conversion", conversionErrorDescription
End Sub

' Conversion font sizes live beside the numbering preference, so they apply
' to every document. Reading is cached; a missing or unreadable file (or an
' older file bridge) means both sizes follow the body text.
Private Sub VTLoadConversionFontSizes()
    Dim encodedSizes As String
    Dim sizesLine As String
    Dim fields As Variant
    Dim sizeField As Variant
    Dim separatorAt As Long
    Dim sizeName As String
    Dim sizeValue As Double

    If VT_WORD_CONVERSION_FONT_SIZES_LOADED Then Exit Sub
    VT_WORD_CONVERSION_FONT_SIZES_LOADED = True
    VT_WORD_INLINE_CONVERSION_FONT_SIZE_PT = 0#
    VT_WORD_DISPLAY_CONVERSION_FONT_SIZE_PT = 0#

    On Error GoTo SizesUnavailable
    encodedSizes = VTFileBridgeCall( _
        "ReadVisualTeXConversionFontSizes", "")
    If Len(encodedSizes) = 0 Then Exit Sub
    sizesLine = VTBase64UrlDecodeUtf8(encodedSizes)
    If Len(sizesLine) = 0 Or Len(sizesLine) > 256 Then Exit Sub
    fields = Split(Trim$(sizesLine), "|")
    For Each sizeField In fields
        separatorAt = InStr(1, CStr(sizeField), "=", vbBinaryCompare)
        If separatorAt > 1 Then
            sizeName = LCase$(Trim$(Left$(CStr(sizeField), separatorAt - 1)))
            ' Val reads "." as the decimal point whatever the locale.
            sizeValue = Val(Trim$(Mid$(CStr(sizeField), separatorAt + 1)))
            If Not VTValidWordFormulaFontSize(sizeValue) Then sizeValue = 0#
            If sizeName = "inline" Then
                VT_WORD_INLINE_CONVERSION_FONT_SIZE_PT = sizeValue
            ElseIf sizeName = "display" Then
                VT_WORD_DISPLAY_CONVERSION_FONT_SIZE_PT = sizeValue
            End If
        End If
    Next sizeField
    Exit Sub

SizesUnavailable:
    VT_WORD_INLINE_CONVERSION_FONT_SIZE_PT = 0#
    VT_WORD_DISPLAY_CONVERSION_FONT_SIZE_PT = 0#
    Err.Clear
End Sub

Public Sub VTWriteConversionFontSizes( _
    ByVal inlineFontSizePt As Double, _
    ByVal displayFontSizePt As Double)

    Dim ignoredResponse As String
    Dim errorNumber As Long
    Dim errorDescription As String

    If Not VTValidWordFormulaFontSize(inlineFontSizePt) Then inlineFontSizePt = 0#
    If Not VTValidWordFormulaFontSize(displayFontSizePt) Then displayFontSizePt = 0#
    On Error GoTo SizesWriteFailed
    ignoredResponse = VTFileBridgeCall( _
        "WriteVisualTeXConversionFontSizes", _
        VTBase64UrlEncodeUtf8( _
            "inline=" & VTJsonNumber(inlineFontSizePt) & _
            "|display=" & VTJsonNumber(displayFontSizePt)))
    VT_WORD_CONVERSION_FONT_SIZES_LOADED = True
    VT_WORD_INLINE_CONVERSION_FONT_SIZE_PT = inlineFontSizePt
    VT_WORD_DISPLAY_CONVERSION_FONT_SIZE_PT = displayFontSizePt
    Exit Sub

SizesWriteFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    Err.Raise errorNumber, "VisualTeX conversion font size", _
        "Unable to save the conversion font size: " & errorDescription
End Sub

' The size set for converting formulas of this display mode ("inline" or
' "block"), or 0 when conversions keep their usual size.
Public Function VTConversionFontSizeSetting( _
    ByVal displayMode As String) As Double

    VTLoadConversionFontSizes
    If displayMode = "block" Then
        VTConversionFontSizeSetting = VT_WORD_DISPLAY_CONVERSION_FONT_SIZE_PT
    Else
        VTConversionFontSizeSetting = VT_WORD_INLINE_CONVERSION_FONT_SIZE_PT
    End If
End Function

' The size a conversion gives a formula: the global setting for its display
' mode when one is set, otherwise the size the conversion found.
Public Function VTConversionFontSize( _
    ByVal displayMode As String, _
    ByVal foundFontSizePt As Double) As Double

    Dim settingPt As Double
    settingPt = VTConversionFontSizeSetting(displayMode)
    If VTValidWordFormulaFontSize(settingPt) Then
        VTConversionFontSize = settingPt
    Else
        VTConversionFontSize = foundFontSizePt
    End If
End Function

' Adds the conversion sizes to a document-import request so the import window
' can size its inline and display formulas. Absent sizes follow the text.
Public Function VTWithConversionFontSizes(ByVal requestJson As String) As String
    Dim inlineSettingPt As Double
    Dim displaySettingPt As Double
    Dim extraFields As String

    inlineSettingPt = VTConversionFontSizeSetting("inline")
    displaySettingPt = VTConversionFontSizeSetting("block")
    If VTValidWordFormulaFontSize(inlineSettingPt) Then
        extraFields = extraFields & ",""inlineFontSizePt"":" & VTJsonNumber(inlineSettingPt)
    End If
    If VTValidWordFormulaFontSize(displaySettingPt) Then
        extraFields = extraFields & ",""displayFontSizePt"":" & VTJsonNumber(displaySettingPt)
    End If
    If Len(extraFields) = 0 Or Right$(requestJson, 2) <> "}}" Then
        VTWithConversionFontSizes = requestJson
        Exit Function
    End If
    ' The request ends with its documentImport object.
    VTWithConversionFontSizes = _
        Left$(requestJson, Len(requestJson) - 2) & extraFields & "}}"
End Function

Public Function VTConversionFontSizeControlMode( _
    ByVal control As IRibbonControl) As String

    If control.Id = VT_WORD_DISPLAY_CONVERSION_FONT_CONTROL_ID Then
        VTConversionFontSizeControlMode = "block"
    Else
        VTConversionFontSizeControlMode = "inline"
    End If
End Function

Public Sub VTRetireLegacyNativeBoundary(ByVal exact As Range, ByVal formulaId As String)
    Dim doc As Document
    Dim owner As Bookmark
    Dim boundary As Range
    Dim following As Range
    Dim nextMath As OMath
    Dim ownerName As String
    Set doc = exact.Document
    ownerName = VTNativeFormulaBookmarkName(formulaId)
    If Not doc.Bookmarks.Exists(ownerName) Then Exit Sub
    Set owner = doc.Bookmarks(ownerName)
    If owner.Range.Start < exact.Start Or owner.Range.End > exact.End Then Exit Sub
    If exact.End >= doc.Content.End - 1 Then Exit Sub
    Set boundary = doc.Range(exact.End, exact.End + 1)
    ' Former native inserts placed this character immediately after their
    ' private owner bookmark. Retire only that proven legacy boundary before
    ' deleting its ownership; ordinary document text is never searched globally.
    If boundary.Text <> ChrW(8288) Or boundary.OMaths.Count <> 0 Then Exit Sub
    If boundary.End < doc.Content.End - 1 Then
        Set following = doc.Range(boundary.End, boundary.End + 1)
        If following.OMaths.Count = 1 Then
            Set nextMath = VTNativeMathRoot(following.OMaths(1))
            If nextMath.Range.Start = boundary.End Then
                ' Preserve two existing native zones with ordinary Word text.
                ' Removing their sole separator would join an untouched formula.
                boundary.Text = " "
                Exit Sub
            End If
        End If
    End If
    boundary.Delete
End Sub

' --- Update Equation Format (one-click upgrade of older Mac formats) ---------

' Older VisualTeX tags this add-in itself wrote around Word equations: the
' VT_F_ identity of a tagged native equation (with its payload variables) and
' the retired numbering-style variables.
Public Function VTV2LegacyTagCount(ByVal doc As Document) As Long
    Dim mark As Bookmark
    Dim variable As Variable
    Dim formulaId As String
    Dim hidden As Boolean

    hidden = doc.Bookmarks.ShowHidden
    doc.Bookmarks.ShowHidden = True
    For Each mark In doc.Bookmarks
        If VTTryFormulaIdFromNativeBookmark(mark.Name, formulaId) Then _
            VTV2LegacyTagCount = VTV2LegacyTagCount + 1
    Next mark
    doc.Bookmarks.ShowHidden = hidden
    For Each variable In doc.Variables
        If variable.Name = VT_WORD_NUMBERING_MODE_VARIABLE Or _
           variable.Name = VT_WORD_NUMBERING_SEPARATOR_VARIABLE Then _
            VTV2LegacyTagCount = VTV2LegacyTagCount + 1
    Next variable
End Function

' The numbered picture formulas in the document's number format: each one's
' number written again in its paragraph (VTWriteSingleParagraphImageNumber)
' with the chapter prefix and restart of its place, every prefix from one
' heading pass. keepCurrent: a picture whose number already is exactly that
' keeps everything (renumbering, opening a document); writing one costs about
' 60 ms in Word for Mac and adds to Undo. Returns how many were written.
Public Function VTUpgradeImageNumberFormats( _
    ByVal doc As Document, _
    Optional ByVal keepCurrent As Boolean = False) As Long
    Dim index As Long
    Dim shape As InlineShape
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim kept As Boolean
    Dim result As Range
    Dim imageShapes As New Collection
    Dim numberingMode As String
    Dim imageCount As Long
    Dim imageAnchors() As Long
    Dim imagePrefixes() As String
    Dim imageOrdinals() As Long

    numberingMode = VTEquationNumberingMode(doc)
    For index = 1 To doc.InlineShapes.Count
        Set shape = doc.InlineShapes(index)
        If VTTryParseFormulaReference(shape.Title, formulaId, displayMode, numbered) Then
            If numbered And displayMode = "block" Then
                imageCount = imageCount + 1
                imageShapes.Add shape
                If imageCount = 1 Then
                    ReDim imageAnchors(1 To 1)
                Else
                    ReDim Preserve imageAnchors(1 To imageCount)
                End If
                imageAnchors(imageCount) = shape.Range.Start
            End If
        End If
    Next index
    If imageCount > 0 And _
       numberingMode <> VT_WORD_NUMBERING_MODE_SEQUENCE Then
        VTBuildEquationHeadingPrefixes _
            doc, imageAnchors, imageCount, numberingMode, _
            imagePrefixes, imageOrdinals
    End If
    For index = imageCount To 1 Step -1
        Set shape = imageShapes(index)
        If Not VTTryParseFormulaReference( _
           shape.Title, formulaId, displayMode, numbered) Then
            Err.Raise vbObjectError + 7605, "VisualTeX", _
                "A numbered image identity changed during format migration."
        End If
        kept = False
        If VTRangeIsWithinTable(shape.Range) Then
            Set result = VTInsertEquationNumber( _
                shape, formulaId, "", True)
        ElseIf numberingMode = VT_WORD_NUMBERING_MODE_SEQUENCE Then
            Set result = VTWriteSingleParagraphImageNumber( _
                shape, formulaId, "", True, False, True, _
                keepCurrent:=keepCurrent, kept:=kept)
        Else
            Set result = VTWriteSingleParagraphImageNumber( _
                shape, formulaId, imagePrefixes(index), True, _
                imageOrdinals(index) = 1, True, _
                keepCurrent:=keepCurrent, kept:=kept)
        End If
        If Not kept Then _
            VTUpgradeImageNumberFormats = VTUpgradeImageNumberFormats + 1
    Next index
    VTWordPerformanceNote "picturesWritten", CStr(VTUpgradeImageNumberFormats)
End Function

Public Sub VTUpgradeWordNumberFormats( _
    ByVal doc As Document, _
    Optional ByVal keepCurrent As Boolean = False)
    Dim formulaId As String
    Dim result As Range
    Dim math As OMath
    Dim bookmark As Bookmark
    Dim ids As New Collection
    Dim item As Variant

    VTPinEquationNumbering doc
    VTUpgradeImageNumberFormats doc, keepCurrent
    VTWordPerformanceMark "formats-pictures"

    For Each bookmark In doc.Bookmarks
        If VTTryFormulaIdFromNativeBookmark(bookmark.Name, formulaId) Then ids.Add formulaId
    Next bookmark
    For Each item In ids
        formulaId = CStr(item)
        Set math = VTNativeMathForBookmark(doc.Bookmarks(VTNativeFormulaBookmarkName(formulaId)))
        If Not math Is Nothing Then
            Set result = math.Range.Duplicate
            If Not result.Information(wdWithInTable) Then
                If Not VTNativeEquationArrayReferenceField(result, formulaId) Is Nothing Then
                    Set result = VTWritePureNativeNumber(result, formulaId)
                End If
                VTFinishPureNativeEquation result, formulaId
            End If
        End If
    Next item
    ' Equation format v2: every native number in the document's numbering
    ' style, written by the shared builder (Windows PlanFormat).
    VTV2RestyleNumbers doc
    VTWordPerformanceMark "formats-native"
End Sub
