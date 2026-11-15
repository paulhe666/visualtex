Attribute VB_Name = "VTWordImages"
Option Explicit
Option Private Module

#Const VT_WORD_DIAGNOSTICS = False

' Image formulas: the picture and its MACROBUTTON container, owner bookmarks, copies, size, baseline, ink
' centre, letter font and typography, and the size watcher.

' A change in Word's font size box reaches a selected image formula this soon.
Private Const VT_WORD_SELECTED_IMAGE_WATCH_SECONDS As Long = 1
Private Const VT_WORD_IMAGE_SCALE_VARIABLE_PREFIX As String = "VT_ImageScale_"
Private Const VT_WORD_IMAGE_CANVAS_VARIABLE_PREFIX As String = "VT_ImageCanvas_"
Private Const VT_WORD_IMAGE_INK_CENTER_VARIABLE_PREFIX As String = _
    "VT_ImageInkCenter_"
Private Const VT_WORD_IMAGE_LETTER_FONT_VARIABLE_PREFIX As String = _
    "VT_ImageLetterFont_"
' Word InlineShape.Font.Position is integral, while the SVG mathematical
' baseline is fractional. Keep the fractional 14 pt reference and map it to the
' final image height immediately before each Word baseline write.
Private Const VT_WORD_MIN_FORMULA_FONT_SIZE_PT As Double = 1#
Private Const VT_WORD_IMAGE_EDIT_MACRO As String = _
    "VisualTeX_EditImageField"
Private VT_WORD_LAST_IMAGE_EDIT_FORMULA_ID As String
Private VT_WORD_LAST_IMAGE_EDIT_AT As Single

Public Function VTIsVisualTeXImageMacroButtonField( _
    ByVal candidateField As Field) As Boolean

    Dim fieldCode As String
    Dim expectedPrefix As String

    If candidateField Is Nothing Then Exit Function
    If candidateField.Type <> wdFieldMacroButton Then Exit Function
    fieldCode = UCase$(VTNormalizedWordFieldCode(candidateField.Code.Text))
    expectedPrefix = UCase$("MACROBUTTON " & VT_WORD_IMAGE_EDIT_MACRO)
    If Left$(fieldCode, Len(expectedPrefix)) <> expectedPrefix Then Exit Function
    If Len(fieldCode) > Len(expectedPrefix) Then
        If Mid$(fieldCode, Len(expectedPrefix) + 1, 1) <> " " Then Exit Function
    End If
    VTIsVisualTeXImageMacroButtonField = True
End Function

Public Function VTContainingFieldForInlineShape( _
    ByVal formulaShape As InlineShape) As Field

    Dim paragraphRange As Range
    Dim candidateField As Field
    Dim match As Field
    Dim matchCount As Long

    If formulaShape Is Nothing Then Exit Function
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    For Each candidateField In paragraphRange.Fields
        On Error Resume Next
        Err.Clear
        If candidateField.Result.Start <= formulaShape.Range.Start And _
           candidateField.Result.End >= formulaShape.Range.End Then
            If Err.Number = 0 Then
                matchCount = matchCount + 1
                Set match = candidateField
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next candidateField
    If matchCount > 1 Then
        Err.Raise vbObjectError + 7593, "VisualTeX", _
            "The VisualTeX image is nested in multiple Word fields."
    End If
    If matchCount = 1 Then Set VTContainingFieldForInlineShape = match
End Function

Public Function VTVisualTeXImageMacroButtonFieldForShape( _
    ByVal formulaShape As InlineShape) As Field

    Dim containingField As Field

    Set containingField = VTContainingFieldForInlineShape(formulaShape)
    If Not containingField Is Nothing Then
        If VTIsVisualTeXImageMacroButtonField(containingField) Then
            Set VTVisualTeXImageMacroButtonFieldForShape = containingField
        End If
    End If
End Function

Public Function VTVisualTeXImageMacroButtonFieldNearPosition( _
    ByVal documentObject As Document, _
    ByVal expectedStart As Long, _
    Optional ByVal maximumDistance As Long = 12) As Field

    Dim candidateField As Field
    Dim match As Field
    Dim candidateStart As Long
    Dim candidateDistance As Long
    Dim bestDistance As Long
    Dim matchCount As Long

    If documentObject Is Nothing Or expectedStart < 0 Or _
       maximumDistance < 0 Then Exit Function
    bestDistance = 2147483647
    For Each candidateField In documentObject.Fields
        On Error Resume Next
        Err.Clear
        If VTIsVisualTeXImageMacroButtonField(candidateField) Then
            candidateStart = candidateField.Code.Start - 1
            candidateDistance = Abs(candidateStart - expectedStart)
            If Err.Number = 0 And candidateDistance <= maximumDistance Then
                If candidateDistance < bestDistance Then
                    bestDistance = candidateDistance
                    matchCount = 1
                    Set match = candidateField
                ElseIf candidateDistance = bestDistance Then
                    matchCount = matchCount + 1
                End If
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next candidateField
    If matchCount = 1 Then
        Set VTVisualTeXImageMacroButtonFieldNearPosition = match
    End If
End Function

Public Function VTVisualTeXImageContainerRange( _
    ByVal formulaShape As InlineShape) As Range

    Dim macroField As Field

    If formulaShape Is Nothing Then Exit Function
    Set macroField = VTVisualTeXImageMacroButtonFieldForShape(formulaShape)
    If macroField Is Nothing Then
        Set VTVisualTeXImageContainerRange = formulaShape.Range.Duplicate
    Else
        Set VTVisualTeXImageContainerRange = macroField.Result.Duplicate
        VTVisualTeXImageContainerRange.Start = macroField.Code.Start - 1
        VTVisualTeXImageContainerRange.End = macroField.Result.End + 1
    End If
End Function

Public Function VTMacroButtonResultHasOnlyImage( _
    ByVal macroField As Field) As Boolean

    Dim resultShape As InlineShape

    If macroField Is Nothing Or _
       Not VTIsVisualTeXImageMacroButtonField(macroField) Then Exit Function
    If macroField.Result.InlineShapes.Count <> 1 Or _
       macroField.Result.OMaths.Count <> 0 Then Exit Function
    Set resultShape = macroField.Result.InlineShapes(1)
    If macroField.Result.Start <> resultShape.Range.Start Or _
       macroField.Result.End <> resultShape.Range.End Then Exit Function
    VTMacroButtonResultHasOnlyImage = True
End Function

Public Function VTParagraphHasSingleVisualTeXImageMacroButton( _
    ByVal paragraphRange As Range) As Boolean

    Dim formulaShape As InlineShape
    Dim formulaContainerRange As Range
    Dim candidateField As Field
    Dim targetBookmarkName As String
    Dim managedFieldCount As Long

    ' Compatibility name retained for existing call sites. The current image is
    ' always one plain InlineShape with no enclosing Word field. A numbered image
    ' may additionally own exactly one visible REF to a VisualTeX VT_N_ sequence
    ' Bookmark after the image; the true SEQ lives in its adjacent helper paragraph.
    If paragraphRange Is Nothing Then Exit Function
    If paragraphRange.InlineShapes.Count <> 1 Or _
       paragraphRange.OMaths.Count <> 0 Then Exit Function
    Set formulaShape = paragraphRange.InlineShapes(1)
    If Not VTIsVisualTeXInlineShape(formulaShape) Then Exit Function
    If Not VTContainingFieldForInlineShape(formulaShape) Is Nothing Then Exit Function
    Set formulaContainerRange = VTVisualTeXImageContainerRange(formulaShape)
    For Each candidateField In paragraphRange.Fields
        If candidateField.Type <> wdFieldRef Or _
           VTEquationFieldStart(candidateField) <= formulaContainerRange.End Then
            Exit Function
        End If
        targetBookmarkName = VTReferenceTargetBookmarkName( _
            candidateField.Code.Text)
        If Left$(targetBookmarkName, _
           Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)) <> _
           VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX Then Exit Function
        managedFieldCount = managedFieldCount + 1
        If managedFieldCount > 1 Then Exit Function
    Next candidateField
    VTParagraphHasSingleVisualTeXImageMacroButton = True
End Function

Public Sub VTDeleteVisualTeXImageContainer( _
    ByVal formulaShape As InlineShape)

    Dim macroField As Field

    If formulaShape Is Nothing Then Exit Sub
    Set macroField = VTVisualTeXImageMacroButtonFieldForShape(formulaShape)
    If macroField Is Nothing Then
        formulaShape.Delete
    Else
        macroField.Delete
    End If
End Sub

Public Sub VTCleanupDocumentEmptyVisualTeXImageMacroButtons( _
    ByVal documentObject As Document)

    Dim candidateField As Field
    Dim fieldIndex As Long

    If documentObject Is Nothing Then Exit Sub
    For fieldIndex = documentObject.Fields.Count To 1 Step -1
        Set candidateField = documentObject.Fields(fieldIndex)
        If VTIsVisualTeXImageMacroButtonField(candidateField) Then
            If candidateField.Result.InlineShapes.Count = 0 Then
                candidateField.Delete
            End If
        End If
    Next fieldIndex
End Sub

Public Function VTDocumentImageMacroSchemaCurrent( _
    ByVal documentObject As Document) As Boolean

    Dim schemaValue As String

    If documentObject Is Nothing Then Exit Function
    If VTTryGetDocumentVariable( _
       documentObject, VT_WORD_IMAGE_MACRO_SCHEMA_VARIABLE, _
       schemaValue) Then
        VTDocumentImageMacroSchemaCurrent = _
            (schemaValue = VT_WORD_IMAGE_MACRO_SCHEMA_VERSION)
    End If
End Function

Public Function VTInlineShapeHasVisualTeXMetadata( _
    ByVal candidate As InlineShape) As Boolean

    On Error GoTo InvalidShape
    If candidate Is Nothing Then Exit Function
    VTInlineShapeHasVisualTeXMetadata = _
        VTIsEncodedMetadata(candidate.AlternativeText)
    Exit Function

InvalidShape:
    VTInlineShapeHasVisualTeXMetadata = False
End Function

Public Function VTInterceptVisualTeXPictureFormatCommand( _
    ByVal commandName As String) As Boolean

    Dim selectedShape As InlineShape
    Dim editErrorNumber As Long
    Dim editErrorDescription As String

    On Error GoTo VisualTeXEditFailed
    If Not VTTryVisualTeXMetadataShapeAtSelection( _
       Selection, selectedShape) Then Exit Function

    VTBeginWordDoubleClickTraceContext
    VTTraceWordDoubleClick _
        commandName & "-visualtex-metadata", Selection, _
        VTInlineShapeTraceSummary(selectedShape)
    VTRequireWritableWordDocument
    VTWordEditInlineShape selectedShape
    VTEndWordDoubleClickTraceContext
    VTInterceptVisualTeXPictureFormatCommand = True
    Exit Function

VisualTeXEditFailed:
    editErrorNumber = Err.Number
    editErrorDescription = Err.Description
    VTTraceWordDoubleClick _
        commandName & "-visualtex-error", Selection, _
        "error=" & CStr(editErrorNumber) & ":" & editErrorDescription
    VTEndWordDoubleClickTraceContext
    VTShowError _
        "Word image formula edit", editErrorNumber, editErrorDescription
    VTInterceptVisualTeXPictureFormatCommand = True
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
    encodedMetadata = selectedShape.AlternativeText
    If Not VTIsEncodedMetadata(encodedMetadata) Then Exit Function

    Set documentObject = selectedShape.Range.Document
    If VTTryFormulaIdForMetadataPayload( _
       documentObject, encodedMetadata, formulaId) Then
        If Not VTTryReadWordFormulaFormat( _
           documentObject, formulaId, displayMode, numbered) Then Exit Function
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
    End If

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

Public Function VTTryResolveVisualTeXInlineShapeReference( _
    ByVal selectedShape As InlineShape, _
    ByRef formulaId As String, _
    ByRef displayMode As String, _
    ByRef numbered As Boolean, _
    ByRef encodedMetadata As String, _
    ByRef metadataNeedsWrite As Boolean, _
    ByRef formatNeedsWrite As Boolean, _
    Optional ByVal repairShapeProperties As Boolean = True) As Boolean

    Dim documentObject As Document
    Dim storedMetadata As String
    Dim storedDisplayMode As String
    Dim storedNumbered As Boolean
    Dim metadataReadable As Boolean
    Dim formatReadable As Boolean
    Dim referenceReadable As Boolean
    Dim titleValue As String
    Dim alternativeValue As String

    On Error GoTo InvalidShape
    formulaId = ""
    displayMode = ""
    numbered = False
    encodedMetadata = ""
    metadataNeedsWrite = False
    formatNeedsWrite = False
    If selectedShape Is Nothing Then Exit Function
    titleValue = selectedShape.Title
    alternativeValue = selectedShape.AlternativeText

    ' The normal committed representation is self-contained: a compact,
    ' validated formula reference in Title and compressed metadata in
    ' AlternativeText. Avoid re-reading chunked document Variables and writing
    ' the same shape properties on every double-click; keep the slower recovery
    ' path below only for legacy, swapped or damaged Office objects.
    If VTIsEncodedMetadata(alternativeValue) Then
        If VTTryParseFormulaReference( _
           titleValue, formulaId, displayMode, numbered) Then
            encodedMetadata = alternativeValue
            VTTryResolveVisualTeXInlineShapeReference = True
            Exit Function
        End If
    End If

    Set documentObject = selectedShape.Range.Document
    If VTIsEncodedMetadata(alternativeValue) Then
        encodedMetadata = alternativeValue
    ElseIf VTIsEncodedMetadata(titleValue) Then
        ' Word can swap Title and AlternativeText while round-tripping an SVG
        ' drawing. Treat both properties as recoverable carriers, then restore
        ' the canonical VisualTeX layout below.
        encodedMetadata = titleValue
    End If

    referenceReadable = VTTryParseFormulaReference( _
        titleValue, formulaId, displayMode, numbered)
    If Not referenceReadable Then
        referenceReadable = VTTryParseFormulaReference( _
            alternativeValue, formulaId, displayMode, numbered)
    End If
    If referenceReadable Then
        On Error Resume Next
        Err.Clear
        metadataReadable = VTTryReadWordMetadataPayload( _
            documentObject, formulaId, storedMetadata)
        Err.Clear
        On Error GoTo InvalidShape
        If metadataReadable Then
            If Len(encodedMetadata) > 0 And _
               StrComp(storedMetadata, encodedMetadata, _
                   vbBinaryCompare) <> 0 Then GoTo InvalidShape
            encodedMetadata = storedMetadata
        ElseIf Len(encodedMetadata) = 0 Then
            Exit Function
        Else
            ' Legacy image formulas can have a valid compact reference before
            ' their document Variables are populated by the first edit.
            metadataNeedsWrite = True
        End If

        On Error Resume Next
        Err.Clear
        formatReadable = VTTryReadWordFormulaFormat( _
            documentObject, formulaId, storedDisplayMode, storedNumbered)
        Err.Clear
        On Error GoTo InvalidShape
        If formatReadable Then
            displayMode = storedDisplayMode
            numbered = storedNumbered
        Else
            formatNeedsWrite = True
        End If
        If repairShapeProperties Then
            selectedShape.AlternativeText = encodedMetadata
            selectedShape.Title = _
                VTFormulaReference(formulaId, displayMode, numbered)
            If StrComp(selectedShape.AlternativeText, encodedMetadata, _
               vbBinaryCompare) <> 0 Then GoTo InvalidShape
        End If
        VTTryResolveVisualTeXInlineShapeReference = True
        Exit Function
    End If

    If Len(encodedMetadata) = 0 Then Exit Function
    If Not repairShapeProperties Then
        ' Ribbon and fingerprint queries must never repair a shape. A unique
        ' persisted payload can still identify legacy images without writing.
        If Not VTTryFormulaIdForMetadataPayload( _
           documentObject, encodedMetadata, formulaId) Then Exit Function
        VTTryResolveVisualTeXInlineShapeReference = _
            VTTryReadWordFormulaFormat( _
                documentObject, formulaId, displayMode, numbered)
        Exit Function
    End If
    ' Reference-free legacy/batch images are resolved by matching their
    ' compressed payload to the document Variables, then both shape properties
    ' are repaired before the edit Session is launched.
    selectedShape.AlternativeText = encodedMetadata
    If Not VTTryRestoreVisualTeXInlineShapeReference( _
       selectedShape) Then Exit Function
    If Not VTTryParseFormulaReference( _
       selectedShape.Title, formulaId, displayMode, numbered) Then Exit Function
    encodedMetadata = selectedShape.AlternativeText
    VTTryResolveVisualTeXInlineShapeReference = True
    Exit Function

InvalidShape:
    formulaId = ""
    displayMode = ""
    numbered = False
    encodedMetadata = ""
    metadataNeedsWrite = False
    formatNeedsWrite = False
    VTTryResolveVisualTeXInlineShapeReference = False
End Function

Public Function VTWordImageOwnerBookmarkName( _
    ByVal formulaId As String) As String

    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7601, "VisualTeX", _
            "VisualTeX cannot bookmark an image formula with an invalid formula id."
    End If
    VTWordImageOwnerBookmarkName = _
        VT_WORD_IMAGE_OWNER_BOOKMARK_PREFIX & Replace$(formulaId, "-", "")
    If Len(VTWordImageOwnerBookmarkName) > 40 Then
        Err.Raise vbObjectError + 7602, "VisualTeX", _
            "VisualTeX generated an image formula bookmark longer than Word permits."
    End If
End Function

Public Function VTWordImageOwnerShape( _
    ByVal documentObject As Document, _
    ByVal formulaId As String) As InlineShape

    Dim bookmarkName As String
    Dim ownerRange As Range

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then Exit Function
    bookmarkName = VTWordImageOwnerBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(bookmarkName) Then Exit Function
    Set ownerRange = documentObject.Bookmarks(bookmarkName).Range.Duplicate
    If ownerRange.InlineShapes.Count <> 1 Then Exit Function
    Set VTWordImageOwnerShape = ownerRange.InlineShapes(1)
End Function

Public Sub VTDeleteWordImageOwnerBookmark( _
    ByVal documentObject As Document, _
    ByVal formulaId As String)

    Dim bookmarkName As String

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then Exit Sub
    bookmarkName = VTWordImageOwnerBookmarkName(formulaId)
    On Error Resume Next
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        documentObject.Bookmarks(bookmarkName).Delete
    End If
    On Error GoTo 0
End Sub

Private Function VTWordImageShapeIsOwner( _
    ByVal documentObject As Document, _
    ByVal formulaShape As InlineShape, _
    ByVal formulaId As String) As Boolean

    Dim ownerShape As InlineShape

    If documentObject Is Nothing Or formulaShape Is Nothing Then Exit Function
    Set ownerShape = VTWordImageOwnerShape(documentObject, formulaId)
    If ownerShape Is Nothing Then Exit Function
    VTWordImageShapeIsOwner = _
        ownerShape.Range.Start = formulaShape.Range.Start And _
        ownerShape.Range.End = formulaShape.Range.End
End Function

Public Sub VTSetWordImageOwnerBookmark( _
    ByVal documentObject As Document, _
    ByVal formulaShape As InlineShape, _
    ByVal formulaId As String)

    Dim bookmarkName As String
    Dim ownerRange As Range
    Dim persistedOwner As InlineShape

    If documentObject Is Nothing Or formulaShape Is Nothing Then
        Err.Raise vbObjectError + 7603, "VisualTeX", _
            "VisualTeX cannot persist a missing image formula owner."
    End If
    bookmarkName = VTWordImageOwnerBookmarkName(formulaId)
    Set ownerRange = formulaShape.Range.Duplicate
    On Error Resume Next
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        documentObject.Bookmarks(bookmarkName).Delete
    End If
    On Error GoTo OwnerFailed
    documentObject.Bookmarks.Add Name:=bookmarkName, Range:=ownerRange
    Set persistedOwner = VTWordImageOwnerShape(documentObject, formulaId)
    If persistedOwner Is Nothing Or _
       persistedOwner.Range.Start <> formulaShape.Range.Start Or _
       persistedOwner.Range.End <> formulaShape.Range.End Then
        Err.Raise vbObjectError + 7603, "VisualTeX", _
            "Word did not persist the VisualTeX image formula owner bookmark."
    End If
    Exit Sub

OwnerFailed:
    Err.Raise Err.Number, "VisualTeX image owner bookmark", Err.Description
End Sub

Public Function VTShouldForkWordImageFormulaCopy( _
    ByVal documentObject As Document, _
    ByVal selectedShape As InlineShape, _
    ByVal formulaId As String, _
    ByVal encodedMetadata As String) As Boolean

    Dim candidate As InlineShape
    Dim canonicalShape As InlineShape
    Dim canonicalRange As Range
    Dim candidateFormulaId As String
    Dim candidateDisplayMode As String
    Dim candidateNumbered As Boolean
    Dim storedMetadata As String
    Dim storedDisplayMode As String
    Dim storedNumbered As Boolean
    Dim identityCount As Long

    If documentObject Is Nothing Or selectedShape Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Or _
       Not VTIsEncodedMetadata(encodedMetadata) Then Exit Function

    ' A formula pasted from another Word document carries its shape metadata but
    ' not this document's formula-scoped Variables. Rekey it before it can bind
    ' to the source formula's global cache identity.
    If Not VTTryReadWordMetadataPayload( _
       documentObject, formulaId, storedMetadata) Then
        VTShouldForkWordImageFormulaCopy = True
        Exit Function
    End If
    If StrComp(storedMetadata, encodedMetadata, vbBinaryCompare) <> 0 Then
        VTShouldForkWordImageFormulaCopy = True
        Exit Function
    End If
    If Not VTTryReadWordFormulaFormat( _
       documentObject, formulaId, storedDisplayMode, storedNumbered) Then
        VTShouldForkWordImageFormulaCopy = True
        Exit Function
    End If

    ' Numbered display formulas have a stronger durable identity than the generic
    ' image-owner bookmark: VT_R_/VT_N_/VT_C_ remain attached to the original
    ' numbered equation when Word duplicates only the formula picture. Resolve
    ' that numbered owner first so a copied display image can never become the
    ' canonical source merely because VT_I_ was dropped or moved by Word.
    If storedDisplayMode = "block" And storedNumbered Then
        Set canonicalRange = VTNumberedFormulaRangeForId( _
            documentObject, formulaId)
        If Not canonicalRange Is Nothing Then
            If canonicalRange.InlineShapes.Count = 1 And _
               canonicalRange.OMaths.Count = 0 Then
                Set canonicalShape = canonicalRange.InlineShapes(1)
                If canonicalShape.Range.Start = selectedShape.Range.Start And _
                   canonicalShape.Range.End = selectedShape.Range.End Then
                    If Not VTWordImageShapeIsOwner( _
                       documentObject, canonicalShape, formulaId) Then
                        VTSetWordImageOwnerBookmark _
                            documentObject, canonicalShape, formulaId
                    End If
                    VTShouldForkWordImageFormulaCopy = False
                Else
                    VTShouldForkWordImageFormulaCopy = True
                End If
                Exit Function
            End If
        End If
    End If

    ' Current formulas have one durable physical owner bookmark. Word does not
    ' clone a same-document bookmark name onto a pasted copy, so the original
    ' remains canonical while only the copied InlineShape is forked.
    Set canonicalShape = VTWordImageOwnerShape(documentObject, formulaId)
    If Not canonicalShape Is Nothing Then
        VTShouldForkWordImageFormulaCopy = _
            Not VTWordImageShapeIsOwner( _
                documentObject, selectedShape, formulaId)
        Exit Function
    End If

    ' Migrate pre-owner-bookmark formulas lazily. With one physical object the
    ' selected formula is necessarily the canonical legacy object. If several
    ' identical legacy objects already exist, preserve the first in document
    ' order and fork only the others when they are edited.
    For Each candidate In documentObject.InlineShapes
        candidateFormulaId = ""
        candidateDisplayMode = ""
        candidateNumbered = False
        On Error Resume Next
        If VTTryParseFormulaReference( _
           candidate.Title, candidateFormulaId, _
           candidateDisplayMode, candidateNumbered) Then
            If candidateFormulaId = formulaId Then
                identityCount = identityCount + 1
                If canonicalShape Is Nothing Then Set canonicalShape = candidate
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next candidate
    If canonicalShape Is Nothing Then
        VTShouldForkWordImageFormulaCopy = True
        Exit Function
    End If
    VTSetWordImageOwnerBookmark documentObject, canonicalShape, formulaId
    VTShouldForkWordImageFormulaCopy = _
        Not VTWordImageShapeIsOwner(documentObject, selectedShape, formulaId)
End Function

Public Function VTIsCopiedWordImageFormula( _
    ByVal selectedShape As InlineShape) As Boolean

    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean
    Dim documentObject As Document

    If selectedShape Is Nothing Then Exit Function
    If Not VTTryResolveVisualTeXInlineShapeReference( _
       selectedShape, formulaId, displayMode, numbered, _
       encodedMetadata, metadataNeedsWrite, formatNeedsWrite) Then Exit Function
    Set documentObject = selectedShape.Range.Document
    VTIsCopiedWordImageFormula = _
        VTShouldForkWordImageFormulaCopy( _
            documentObject, selectedShape, formulaId, encodedMetadata)
End Function

Public Function VTTryResolveVisualTeXInlineShapeAtSelection( _
    ByVal selected As Selection, _
    ByRef resolvedShape As InlineShape, _
    ByRef formulaId As String, _
    ByRef displayMode As String, _
    ByRef numbered As Boolean, _
    ByRef encodedMetadata As String, _
    ByRef metadataNeedsWrite As Boolean, _
    ByRef formatNeedsWrite As Boolean, _
    Optional ByVal repairShapeProperties As Boolean = True) As Boolean

    Dim probeRange As Range
    Dim candidate As InlineShape
    Dim candidateField As Field
    Dim match As InlineShape
    Dim candidateFormulaId As String
    Dim candidateDisplayMode As String
    Dim candidateNumbered As Boolean
    Dim candidateMetadata As String
    Dim candidateMetadataNeedsWrite As Boolean
    Dim candidateFormatNeedsWrite As Boolean
    Dim matchFormulaId As String
    Dim matchDisplayMode As String
    Dim matchNumbered As Boolean
    Dim matchMetadata As String
    Dim matchMetadataNeedsWrite As Boolean
    Dim matchFormatNeedsWrite As Boolean
    Dim matchCount As Long
    Dim directCount As Long
    Dim candidateCount As Long
    Dim fieldStart As Long
    Dim fieldEnd As Long

    On Error GoTo InvalidSelection
    Set resolvedShape = Nothing
    formulaId = ""
    displayMode = ""
    numbered = False
    encodedMetadata = ""
    metadataNeedsWrite = False
    formatNeedsWrite = False
    If selected Is Nothing Then
        VTTraceWordDoubleClick "resolver-skip", selected, "reason=no-selection"
        Exit Function
    End If
    directCount = selected.InlineShapes.Count
    VTTraceWordDoubleClick _
        "resolver-enter", selected, _
        "directInlineShapes=" & CStr(directCount)
    If directCount = 1 Then
        VTTraceWordDoubleClick _
            "resolver-direct-candidate", selected, _
            VTInlineShapeTraceSummary(selected.InlineShapes(1))
        If VTTryResolveVisualTeXInlineShapeReference( _
           selected.InlineShapes(1), formulaId, displayMode, numbered, _
           encodedMetadata, metadataNeedsWrite, formatNeedsWrite, _
           repairShapeProperties) Then
            Set resolvedShape = selected.InlineShapes(1)
            VTTraceWordDoubleClick _
                "resolver-direct-match", selected, _
                VTInlineShapeTraceSummary(resolvedShape)
            VTTryResolveVisualTeXInlineShapeAtSelection = True
            Exit Function
        End If
        VTTraceWordDoubleClick "resolver-direct-rejected", selected, ""
    End If

    ' A MACROBUTTON can invoke its macro with the caret on a field boundary
    ' rather than with its result selected. Resolve that exact field before the
    ' generic adjacent-character fallback, without scanning the document.
    For Each candidateField In selected.Range.Paragraphs(1).Range.Fields
        If VTIsVisualTeXImageMacroButtonField(candidateField) And _
           VTMacroButtonResultHasOnlyImage(candidateField) Then
            fieldStart = candidateField.Code.Start - 1
            fieldEnd = candidateField.Result.End + 1
            If selected.Range.Start <= fieldEnd + 1 And _
               selected.Range.End >= fieldStart - 1 Then
                Set candidate = candidateField.Result.InlineShapes(1)
                If VTTryResolveVisualTeXInlineShapeReference( _
                   candidate, candidateFormulaId, candidateDisplayMode, _
                   candidateNumbered, candidateMetadata, _
                   candidateMetadataNeedsWrite, _
                   candidateFormatNeedsWrite, repairShapeProperties) Then
                    matchCount = matchCount + 1
                    Set match = candidate
                    matchFormulaId = candidateFormulaId
                    matchDisplayMode = candidateDisplayMode
                    matchNumbered = candidateNumbered
                    matchMetadata = candidateMetadata
                    matchMetadataNeedsWrite = _
                        candidateMetadataNeedsWrite
                    matchFormatNeedsWrite = candidateFormatNeedsWrite
                End If
            End If
        End If
    Next candidateField
    If matchCount = 1 Then
        Set resolvedShape = match
        formulaId = matchFormulaId
        displayMode = matchDisplayMode
        numbered = matchNumbered
        encodedMetadata = matchMetadata
        metadataNeedsWrite = matchMetadataNeedsWrite
        formatNeedsWrite = matchFormatNeedsWrite
        VTTryResolveVisualTeXInlineShapeAtSelection = True
        VTTraceWordDoubleClick _
            "resolver-field-match", selected, _
            VTInlineShapeTraceSummary(match)
        Exit Function
    End If
    matchCount = 0
    Set match = Nothing

    Set probeRange = selected.Range.Duplicate
    probeRange.MoveStart Unit:=wdCharacter, Count:=-1
    probeRange.MoveEnd Unit:=wdCharacter, Count:=1
    candidateCount = probeRange.InlineShapes.Count
    VTTraceWordDoubleClick _
        "resolver-adjacent-probe", selected, _
        "probeStart=" & CStr(probeRange.Start) & _
        " probeEnd=" & CStr(probeRange.End) & _
        " candidates=" & CStr(candidateCount)
    For Each candidate In probeRange.InlineShapes
        VTTraceWordDoubleClick _
            "resolver-adjacent-candidate", selected, _
            VTInlineShapeTraceSummary(candidate)
        If VTTryResolveVisualTeXInlineShapeReference( _
           candidate, candidateFormulaId, candidateDisplayMode, _
           candidateNumbered, candidateMetadata, _
           candidateMetadataNeedsWrite, candidateFormatNeedsWrite, _
           repairShapeProperties) Then
            matchCount = matchCount + 1
            Set match = candidate
            matchFormulaId = candidateFormulaId
            matchDisplayMode = candidateDisplayMode
            matchNumbered = candidateNumbered
            matchMetadata = candidateMetadata
            matchMetadataNeedsWrite = candidateMetadataNeedsWrite
            matchFormatNeedsWrite = candidateFormatNeedsWrite
            VTTraceWordDoubleClick _
                "resolver-adjacent-match", selected, _
                "matchCount=" & CStr(matchCount) & " " & _
                VTInlineShapeTraceSummary(candidate)
        End If
    Next candidate
    If matchCount = 1 Then
        Set resolvedShape = match
        formulaId = matchFormulaId
        displayMode = matchDisplayMode
        numbered = matchNumbered
        encodedMetadata = matchMetadata
        metadataNeedsWrite = matchMetadataNeedsWrite
        formatNeedsWrite = matchFormatNeedsWrite
        VTTryResolveVisualTeXInlineShapeAtSelection = True
        VTTraceWordDoubleClick "resolver-complete", selected, "result=single-match"
    Else
        VTTraceWordDoubleClick _
            "resolver-complete", selected, _
            "result=no-unique-match matchCount=" & CStr(matchCount)
    End If
    Exit Function

InvalidSelection:
    VTTraceWordDoubleClick _
        "resolver-error", selected, _
        "error=" & CStr(Err.Number) & ":" & Err.Description
    Set resolvedShape = Nothing
    formulaId = ""
    displayMode = ""
    numbered = False
    encodedMetadata = ""
    metadataNeedsWrite = False
    formatNeedsWrite = False
End Function

Private Function VTImageEditDispatchDebounced( _
    ByVal formulaId As String) As Boolean

    Dim elapsedSeconds As Double

    If Len(VT_WORD_LAST_IMAGE_EDIT_FORMULA_ID) = 0 Or _
       StrComp(VT_WORD_LAST_IMAGE_EDIT_FORMULA_ID, formulaId, _
           vbBinaryCompare) <> 0 Then Exit Function
    elapsedSeconds = VTTimerElapsedSeconds( _
        VT_WORD_LAST_IMAGE_EDIT_AT, Timer)
    VTImageEditDispatchDebounced = _
        (elapsedSeconds >= 0# And _
         elapsedSeconds <= VT_WORD_IMAGE_EDIT_DEBOUNCE_SECONDS)
End Function

Public Function VTVisualTeXLetterFontFromWordFontName( _
    ByVal fontName As String) As String

    Select Case LCase$(Trim$(fontName))
        Case "katex_math", "katex math", "katex_main", "katex main", _
             "visualtex latin modern math", "latin modern math"
            VTVisualTeXLetterFontFromWordFontName = "katex"
        Case "times new roman", "times", _
             "visualtex tex gyre termes math", "tex gyre termes math"
            VTVisualTeXLetterFontFromWordFontName = "times"
        Case "cambria math", "cambria"
            VTVisualTeXLetterFontFromWordFontName = "cambria"
        Case "stix two math", "stix two text", "stixgeneral", "stix"
            VTVisualTeXLetterFontFromWordFontName = "stix"
        Case "palatino", "palatino linotype", "book antiqua"
            VTVisualTeXLetterFontFromWordFontName = "palatino"
        Case "helvetica neue", "helvetica", "arial"
            VTVisualTeXLetterFontFromWordFontName = "helvetica"
    End Select
End Function

Public Sub VTWordOpenResolvedInlineShape( _
    ByVal selectedShape As InlineShape, _
    ByVal formulaId As String, _
    ByVal displayMode As String, _
    ByVal numbered As Boolean, _
    ByVal encodedMetadata As String, _
    ByVal metadataNeedsWrite As Boolean, _
    ByVal formatNeedsWrite As Boolean, _
    Optional ByVal convertToNative As Boolean = False, _
    Optional ByVal operationName As String = "formula", _
    Optional ByRef launchedSessionId As Variant)

    Dim documentObject As Document
    Dim sessionId As String
    Dim requestJson As String
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim observedWordFontSizePt As Double
    Dim debounceArmed As Boolean
    Dim migrationMutationStarted As Boolean
    Dim migratedNumberRange As Range
    Dim launchTiming As String
    Dim openErrorNumber As Long
    Dim openErrorDescription As String
    Dim openStage As String
    Dim sourceFormulaId As String
    Dim sourceObjectId As String
    Dim forkCopiedFormula As Boolean
    Dim editBookmarkCreated As Boolean

    VTTraceWordDoubleClick _
        "edit-inline-enter", Selection, _
        "convertToNative=" & CStr(convertToNative) & " " & _
        VTInlineShapeTraceSummary(selectedShape)
    If selectedShape Is Nothing Then
        VTTraceWordDoubleClick "edit-inline-error", Selection, "reason=no-shape"
        Err.Raise vbObjectError + 7400, "VisualTeX", "Select exactly one VisualTeX inline formula."
    End If
    If Not VTIsCanonicalUuid(formulaId) Or _
       (displayMode <> "inline" And displayMode <> "block") Or _
       (numbered And displayMode <> "block") Or _
       Not VTIsEncodedMetadata(encodedMetadata) Then
        VTTraceWordDoubleClick _
            "edit-inline-error", Selection, _
            "reason=metadata-unrecoverable " & _
            VTInlineShapeTraceSummary(selectedShape)
        Err.Raise vbObjectError + 7400, "VisualTeX", _
            "The selected image does not contain recoverable VisualTeX edit metadata."
    End If
    sourceFormulaId = formulaId
    On Error GoTo OpenFailed
    openStage = "resolve-document"
    Set documentObject = selectedShape.Range.Document
    If displayMode = "block" And numbered And _
       selectedShape.Range.Information(wdWithInTable) Then
        openStage = "migrate-numbered-image-table"
        VTBeginWordInternalMutation
        migrationMutationStarted = True
        Set migratedNumberRange = _
            VTMigrateLegacyNumberedImageTableToParagraph( _
                selectedShape, formulaId, "VisualTeX formula")
        VTEndWordInternalMutation
        migrationMutationStarted = False
        If migratedNumberRange.InlineShapes.Count <> 1 Then
            Err.Raise vbObjectError + 7569, "VisualTeX", _
                "The migrated numbered image could not be re-resolved before editing."
        End If
        Set selectedShape = migratedNumberRange.InlineShapes(1)
    End If
    forkCopiedFormula = _
        (operationName = "formula" Or operationName = "imageToNative") And _
        VTShouldForkWordImageFormulaCopy( _
            documentObject, selectedShape, sourceFormulaId, encodedMetadata)
    If Not convertToNative And Not forkCopiedFormula And _
       VTImageEditDispatchDebounced(sourceFormulaId) Then
        VTTraceWordDoubleClick _
            "edit-inline-debounced", Selection, _
            "formulaId=" & sourceFormulaId
        Exit Sub
    End If
    If Not convertToNative And Not forkCopiedFormula Then
        VT_WORD_LAST_IMAGE_EDIT_FORMULA_ID = sourceFormulaId
        VT_WORD_LAST_IMAGE_EDIT_AT = Timer
        debounceArmed = True
    End If
    openStage = "prepare-formula-state"
    If operationName = "imageToNative" Then
        VTReadWordImageConversionState _
            selectedShape, formulaId, fontSizePt, referenceWidthPt, _
            referenceHeightPt, referenceBaselinePt, observedWordFontSizePt
        fontSizePt = VTConversionFontSize(displayMode, fontSizePt)
    Else
        VTPrepareWordImageFormulaState _
            selectedShape, formulaId, displayMode, numbered, fontSizePt, _
            referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
            observedWordFontSizePt
    End If

    openStage = "write-metadata"
    If Not forkCopiedFormula And metadataNeedsWrite Then
        VTSetWordMetadataPayload documentObject, formulaId, encodedMetadata
    End If
    openStage = "write-format"
    If Not forkCopiedFormula And formatNeedsWrite Then
        VTSetWordFormulaFormat documentObject, formulaId, displayMode, numbered
    End If

    openStage = "build-request"
    sessionId = VTNewUuidV4()
    VTAddWordEditBookmark selectedShape, sessionId
    editBookmarkCreated = True
    sourceObjectId = VTWordEditBookmarkName(sessionId)
    If forkCopiedFormula Then formulaId = VTNewUuidV4()
    requestJson = VTRequestJson( _
        sessionId, _
        VT_WORD_HOST, _
        "edit", _
        formulaId, _
        displayMode, _
        numbered, _
        VTWordDocumentIdentity(), _
        sourceObjectId, _
        encodedMetadata, _
        "", _
        "", _
        convertToNative, _
        fontSizePt, _
        referenceWidthPt, _
        referenceHeightPt, _
        operationName, _
        forkCopiedFormula)
    VTTraceWordDoubleClick _
        "edit-inline-resolved", Selection, _
        "formulaId=" & formulaId & _
        " sourceFormulaId=" & sourceFormulaId & _
        " forkCopied=" & CStr(forkCopiedFormula) & _
        " displayMode=" & displayMode & _
        " numbered=" & CStr(numbered)
    requestJson = VTV2WithDocumentMathFont(requestJson, documentObject)
    openStage = "launch-session"
    If displayMode = "block" Then VTPrepareNativeRollbackDocument documentObject
    launchTiming = _
        VTWriteAndLaunchSession(VT_WORD_HOST, sessionId, requestJson)
    If Not IsMissing(launchedSessionId) Then launchedSessionId = sessionId
    VTTraceWordDoubleClick _
        "edit-inline-editor-launched", Selection, _
        "sessionId=" & sessionId & " " & launchTiming
    Exit Sub

OpenFailed:
    openErrorNumber = Err.Number
    openErrorDescription = Err.Description
    If migrationMutationStarted Then
        VTEndWordInternalMutation
        migrationMutationStarted = False
    End If
    If editBookmarkCreated And Not documentObject Is Nothing And _
       VTIsCanonicalUuid(sessionId) Then
        VTDeleteWordEditBookmark documentObject, sessionId
    End If
    If debounceArmed And _
       StrComp(VT_WORD_LAST_IMAGE_EDIT_FORMULA_ID, sourceFormulaId, _
           vbBinaryCompare) = 0 Then
        VT_WORD_LAST_IMAGE_EDIT_FORMULA_ID = ""
        VT_WORD_LAST_IMAGE_EDIT_AT = 0!
    End If
    Err.Raise openErrorNumber, "VisualTeX Word image edit", _
        openStage & ": " & openErrorDescription
End Sub

Public Function VTDispatchVisualTeXImageEditAtSelection( _
    ByVal selected As Selection) As Boolean

    Dim selectedShape As InlineShape
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean

    If Not VTTryResolveVisualTeXInlineShapeAtSelection( _
       selected, selectedShape, formulaId, displayMode, numbered, _
       encodedMetadata, metadataNeedsWrite, formatNeedsWrite) Then Exit Function
    VTRequireWritableWordDocument
    VTWordOpenResolvedInlineShape _
        selectedShape, formulaId, displayMode, numbered, encodedMetadata, _
        metadataNeedsWrite, formatNeedsWrite, False
    VTDispatchVisualTeXImageEditAtSelection = True
End Function

Public Sub VTWordEditInlineShape( _
    ByVal selectedShape As InlineShape, _
    Optional ByVal convertToNative As Boolean = False, _
    Optional ByVal operationName As String = "formula", _
    Optional ByRef launchedSessionId As Variant)

    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean

    If selectedShape Is Nothing Or _
       Not VTTryResolveVisualTeXInlineShapeReference( _
           selectedShape, formulaId, displayMode, numbered, _
           encodedMetadata, metadataNeedsWrite, formatNeedsWrite) Then
        Err.Raise vbObjectError + 7400, "VisualTeX", _
            "The selected image does not contain recoverable VisualTeX edit metadata."
    End If
    If IsMissing(launchedSessionId) Then
        VTWordOpenResolvedInlineShape _
            selectedShape, formulaId, displayMode, numbered, encodedMetadata, _
            metadataNeedsWrite, formatNeedsWrite, convertToNative, operationName
    Else
        VTWordOpenResolvedInlineShape _
            selectedShape, formulaId, displayMode, numbered, encodedMetadata, _
            metadataNeedsWrite, formatNeedsWrite, convertToNative, operationName, _
            launchedSessionId
    End If
End Sub

Public Function VTAddWordFormulaPicture( _
    ByVal documentObject As Document, _
    ByVal targetRange As Range, _
    ByVal vectorDocumentPath As String, _
    ByVal fallbackImagePath As String, _
    Optional ByVal sharedStagingDocument As Variant, _
    Optional ByVal sharedStagingShapeIndex As Long = 0) As InlineShape

    Dim stagingDocument As Document
    Dim stagingShapeRange As Range
    Dim insertionRange As Range
    Dim candidate As InlineShape
    Dim probeShape As InlineShape
    Dim targetStart As Long
    Dim targetEnd As Long
    Dim matchCount As Long
    Dim vectorErrorNumber As Long
    Dim vectorErrorDescription As String
    Dim fallbackErrorNumber As Long
    Dim fallbackErrorDescription As String
    Dim ownsStagingDocument As Boolean

    If documentObject Is Nothing Or targetRange Is Nothing Or _
       Len(vectorDocumentPath) = 0 Then
        Err.Raise vbObjectError + 7570, "VisualTeX", _
            "The Word formula SVG staging target is invalid."
    End If
    targetStart = targetRange.Start
    targetEnd = targetRange.End

    ' Word for Mac can store and render SVG in DOCX, but its VBA
    ' InlineShapes.AddPicture method rejects a raw .svg file. For a standalone
    ' formula, import the tiny cache DOCX into one empty hidden document with
    ' Range.InsertFile, then transfer only its parsed InlineShape through
    ' FormattedText. Creating an empty document is much cheaper than opening the
    ' cache DOCX itself on every OMML -> image conversion. If InsertFile is not
    ' supported by a particular Word build, retain the reviewed Documents.Open
    ' path as a compatibility fallback before falling back to PNG.
    On Error GoTo VectorFailed
    If IsObject(sharedStagingDocument) Then
        Set stagingDocument = sharedStagingDocument
        If sharedStagingShapeIndex < 1 Or _
           sharedStagingShapeIndex > stagingDocument.InlineShapes.Count Then
            Err.Raise vbObjectError + 7570, "VisualTeX", _
                "The shared Word SVG staging index is invalid."
        End If
    Else
        On Error GoTo DirectStageFailed
        Set stagingDocument = Documents.Add(Visible:=False)
        ownsStagingDocument = True
        Set insertionRange = stagingDocument.Content.Duplicate
        insertionRange.Collapse wdCollapseStart
        insertionRange.InsertFile _
            FileName:=vectorDocumentPath, _
            ConfirmConversions:=False, Link:=False, Attachment:=False
        sharedStagingShapeIndex = 1
        If stagingDocument.InlineShapes.Count <> 1 Or _
           stagingDocument.OMaths.Count <> 0 Or _
           stagingDocument.Tables.Count <> 0 Then
            Err.Raise vbObjectError + 7570, "VisualTeX", _
                "Word did not import exactly one SVG drawing into the fast staging document."
        End If
        GoTo StagingReady

DirectStageFailed:
        On Error Resume Next
        If Not stagingDocument Is Nothing Then
            stagingDocument.Saved = True
            stagingDocument.Close SaveChanges:=wdDoNotSaveChanges
        End If
        Set stagingDocument = Nothing
        Err.Clear
        On Error GoTo VectorFailed
        Set stagingDocument = Documents.Open( _
            FileName:=vectorDocumentPath, _
            ConfirmConversions:=False, _
            ReadOnly:=True, _
            AddToRecentFiles:=False, _
            Visible:=False)
        ownsStagingDocument = True
        sharedStagingShapeIndex = 1
        If stagingDocument.InlineShapes.Count <> 1 Then
            Err.Raise vbObjectError + 7570, "VisualTeX", _
                "Word did not parse exactly one SVG drawing from the staging document."
        End If
    End If

StagingReady:
    Set stagingShapeRange = _
        stagingDocument.InlineShapes(sharedStagingShapeIndex).Range.Duplicate
    documentObject.Activate
    Set insertionRange = documentObject.Range( _
        Start:=targetStart, End:=targetEnd)
    insertionRange.FormattedText = stagingShapeRange.FormattedText
    If ownsStagingDocument Then
        stagingDocument.Saved = True
        stagingDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set stagingDocument = Nothing
        documentObject.Activate
    End If

    Set insertionRange = documentObject.Range( _
        Start:=targetStart, End:=targetStart + 1)
    If insertionRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7570, "VisualTeX", _
            "Word did not transfer exactly one SVG formula drawing."
    End If
    Set candidate = insertionRange.InlineShapes(1)
    Set VTAddWordFormulaPicture = candidate
    Exit Function

VectorFailed:
    vectorErrorNumber = Err.Number
    vectorErrorDescription = Err.Description
    On Error Resume Next
    If ownsStagingDocument And Not stagingDocument Is Nothing Then
        stagingDocument.Close SaveChanges:=wdDoNotSaveChanges
        documentObject.Activate
    End If
    Err.Clear
    On Error GoTo FallbackFailed

    ' The PNG is only an emergency compatibility path. A vector-import failure
    ' must not leave Word's generic "cannot open formula.svg" modal dialog or
    ' abort an otherwise valid formula transaction.
    If Len(fallbackImagePath) = 0 Then
        Err.Raise vbObjectError + 7570, "VisualTeX", _
            "Word could not import the VisualTeX SVG staging document: " & _
            CStr(vectorErrorNumber) & " " & vectorErrorDescription
    End If
    Set insertionRange = documentObject.Range( _
        Start:=targetStart, End:=targetEnd)
    Set candidate = documentObject.InlineShapes.AddPicture( _
        FileName:=fallbackImagePath, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Range:=insertionRange)
    If candidate Is Nothing Then
        Err.Raise vbObjectError + 7570, "VisualTeX", _
            "Word did not create the PNG compatibility formula picture."
    End If
    Set VTAddWordFormulaPicture = candidate
    Exit Function

FallbackFailed:
    fallbackErrorNumber = Err.Number
    fallbackErrorDescription = Err.Description
    Err.Raise vbObjectError + 7570, "VisualTeX Word vector insertion", _
        "SVG staging failed (" & CStr(vectorErrorNumber) & "): " & _
        vectorErrorDescription & "; PNG fallback failed (" & _
        CStr(fallbackErrorNumber) & "): " & fallbackErrorDescription
End Function

Public Sub VTDeleteTrailingNativeImageArtifact( _
    ByVal formulaShape As InlineShape, _
    Optional ByVal expectedOwnedNumber As String = "")

    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim trailingRange As Range
    Dim trailingText As String
    Dim trailingStart As Long
    Dim cleanupIndex As Long
    Dim contentEnd As Long
    Dim terminalSuffixRange As Range
    Dim formulaContainerRange As Range
    Dim candidateField As Field

    If formulaShape Is Nothing Then Exit Sub
    Set documentObject = formulaShape.Range.Document
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    trailingStart = formulaShape.Range.End

    ' A native OMath replacement can leave one or two invisible boundary
    ' characters behind. Remove only Unicode format characters created by the
    ' conversion; ordinary spaces, tabs, paragraph marks and body text are
    ' intentionally preserved.
    For cleanupIndex = 1 To 2
        If trailingStart >= paragraphRange.End - 1 Then Exit For
        Set trailingRange = documentObject.Range( _
            Start:=trailingStart, End:=trailingStart + 1)
        trailingText = trailingRange.Text
        If trailingText = ChrW(8203) Or _
           trailingText = ChrW(8204) Or _
           trailingText = ChrW(8205) Or _
           trailingText = ChrW(8288) Or _
           trailingText = ChrW(65279) Or _
           trailingText = Chr$(11) Then
            trailingRange.Delete
            Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
        Else
            Exit For
        End If
    Next cleanupIndex

    ' Word may consume the U+2060 math boundary while preserving its following
    ' ordinary space. Strip paragraph/end-of-cell markers from the inspection
    ' boundary, then remove the complete suffix only when it contains no body
    ' text, fields, formulas or drawings. A genuine separator before following
    ' body text therefore remains untouched.
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    contentEnd = paragraphRange.End
    Do While contentEnd > trailingStart
        trailingText = documentObject.Range( _
            Start:=contentEnd - 1, End:=contentEnd).Text
        If trailingText = vbCr Or trailingText = Chr$(7) Then
            contentEnd = contentEnd - 1
        Else
            Exit Do
        End If
    Loop
    If trailingStart < contentEnd Then
        Set terminalSuffixRange = documentObject.Range( _
            Start:=trailingStart, End:=contentEnd)
        If terminalSuffixRange.Fields.Count = 0 And _
           terminalSuffixRange.InlineShapes.Count = 0 And _
           terminalSuffixRange.OMaths.Count = 0 And _
           Not VTWordRangeHasMeaningfulText(terminalSuffixRange) Then
            terminalSuffixRange.Delete
        End If
    End If

    If Len(expectedOwnedNumber) = 0 Then Exit Sub
    Set formulaContainerRange = VTVisualTeXImageContainerRange(formulaShape)
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    contentEnd = paragraphRange.End - 1
    If formulaContainerRange.End >= contentEnd Then Exit Sub
    Set terminalSuffixRange = documentObject.Range( _
        formulaContainerRange.End, contentEnd)
    trailingText = terminalSuffixRange.Text
    For Each candidateField In terminalSuffixRange.Fields
        If candidateField.Type <> wdFieldSequence And _
           candidateField.Type <> wdFieldQuote And _
           candidateField.Type <> wdFieldStyleRef Then Exit Sub
        trailingText = Replace$( _
            trailingText, candidateField.Result.Text, "", 1, -1, _
            vbBinaryCompare)
    Next candidateField
    trailingText = Replace$(trailingText, expectedOwnedNumber, "")
    trailingText = Replace$(trailingText, "#", "")
    trailingText = Replace$(trailingText, "(", "")
    trailingText = Replace$(trailingText, ")", "")
    trailingText = Replace$(trailingText, vbTab, "")
    trailingText = Replace$(trailingText, " ", "")
    trailingText = Replace$(trailingText, ChrW(160), "")
    trailingText = Replace$(trailingText, ChrW(8203), "")
    trailingText = Replace$(trailingText, ChrW(8204), "")
    trailingText = Replace$(trailingText, ChrW(8205), "")
    trailingText = Replace$(trailingText, ChrW(8288), "")
    trailingText = Replace$(trailingText, ChrW(65279), "")
    trailingText = Replace$(trailingText, Chr$(11), "")
    If Len(trailingText) = 0 Then terminalSuffixRange.Delete
End Sub

Public Function VTDetachWordFormulaPictureFromMath( _
    ByVal formulaShape As InlineShape, _
    Optional ByRef targetDocumentMutated As Boolean = False) As InlineShape

    Dim documentObject As Document
    Dim sourceParagraph As Range
    Dim backupDocument As Document
    Dim backupRange As Range
    Dim parsedDocument As Document
    Dim replacementRange As Range
    Dim sourceXml As String
    Dim pictureOrdinal As Long
    Dim pictureIndex As Long
    Dim bodyStart As Long
    Dim bodyEnd As Long
    Dim scanPosition As Long
    Dim pictureCount As Long
    Dim originalTitle As String
    Dim originalMetadata As String
    Dim drawingXml As String
    Dim drawingStart As Long
    Dim drawingEnd As Long
    Dim mathStart As Long
    Dim mathEnd As Long
    Dim mathParagraphStart As Long
    Dim mathParagraphEnd As Long
    Dim paragraphStart As Long
    Dim paragraphEnd As Long
    Dim originalContentEnd As Long
    Dim replacementEnd As Long
    Dim stagingPath As String
    Dim sourceReplaced As Boolean
    Dim internalMutationStarted As Boolean
    Dim detachErrorNumber As Long
    Dim detachErrorDescription As String

    If formulaShape Is Nothing Then
        Err.Raise vbObjectError + 7594, "VisualTeX", _
            "The native-to-image detachment target is missing."
    End If
    If formulaShape.Range.OMaths.Count = 0 Then
        VTDeleteTrailingNativeImageArtifact formulaShape
        Set VTDetachWordFormulaPictureFromMath = formulaShape
        Exit Function
    End If
    If formulaShape.Range.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7594, "VisualTeX", _
            "The converted image is inside an ambiguous Word math container."
    End If

    On Error GoTo DetachFailed
    VTBeginWordInternalMutation
    internalMutationStarted = True
    Set documentObject = formulaShape.Range.Document
    Set sourceParagraph = formulaShape.Range.Paragraphs(1).Range.Duplicate
    ' Preserve an entire paragraph boundary: replacing just the picture or
    ' OMath Range inherits the old math zone on Word for Mac, even after Remove.
    If sourceParagraph.Information(wdWithInTable) Then
        If Right$(sourceParagraph.Text, 1) = Chr$(7) Then _
            sourceParagraph.End = sourceParagraph.End - 1
    End If
    paragraphStart = sourceParagraph.Start
    paragraphEnd = sourceParagraph.End
    originalContentEnd = documentObject.Content.End
    sourceXml = sourceParagraph.WordOpenXML
    originalTitle = formulaShape.Title
    originalMetadata = formulaShape.AlternativeText
    For pictureIndex = 1 To sourceParagraph.InlineShapes.Count
        If sourceParagraph.InlineShapes(pictureIndex).Range.Start = _
           formulaShape.Range.Start Then pictureOrdinal = pictureIndex
    Next pictureIndex
    If pictureOrdinal = 0 Then GoTo InvalidCarrier
    ' Use one export only: separate Range exports can assign different rIds.
    ' Restrict scanning to the main document body, excluding headers and parts.
    bodyStart = InStr(1, sourceXml, "<w:body>", vbBinaryCompare)
    If bodyStart = 0 Then GoTo InvalidCarrier
    bodyEnd = InStr(bodyStart, sourceXml, "</w:body>", vbBinaryCompare)
    If bodyEnd = 0 Then GoTo InvalidCarrier
    scanPosition = bodyStart
    Do
        scanPosition = InStr(scanPosition, sourceXml, "<w:drawing>", vbBinaryCompare)
        If scanPosition = 0 Or scanPosition >= bodyEnd Then Exit Do
        pictureCount = pictureCount + 1
        If pictureCount = pictureOrdinal Then drawingStart = scanPosition
        scanPosition = scanPosition + Len("<w:drawing>")
    Loop
    ' VML/floating drawings do not have a one-to-one InlineShapes order.
    ' Reject them before mutation rather than guessing a target.
    If pictureCount <> sourceParagraph.InlineShapes.Count Then GoTo InvalidCarrier
    If InStr(bodyStart, sourceXml, "<wp:anchor", vbBinaryCompare) > 0 And _
       InStr(bodyStart, sourceXml, "<wp:anchor", vbBinaryCompare) < bodyEnd Then GoTo InvalidCarrier
    If drawingStart = 0 Then GoTo InvalidCarrier
    drawingEnd = InStr(drawingStart, sourceXml, "</w:drawing>", vbBinaryCompare)
    If drawingEnd = 0 Or drawingEnd >= bodyEnd Then GoTo InvalidCarrier
    drawingXml = Mid$(sourceXml, drawingStart, _
        drawingEnd + Len("</w:drawing>") - drawingStart)
    mathStart = InStrRev(sourceXml, "<m:oMath>", drawingStart, vbBinaryCompare)
    mathEnd = InStr(drawingStart, sourceXml, "</m:oMath>", vbBinaryCompare)
    If mathStart = 0 Or mathEnd = 0 Then GoTo InvalidCarrier
    mathEnd = mathEnd + Len("</m:oMath>")
    mathParagraphStart = InStrRev(sourceXml, "<m:oMathPara>", mathStart, vbBinaryCompare)
    If mathParagraphStart > 0 Then
        mathParagraphEnd = InStr(mathParagraphStart, sourceXml, _
            "</m:oMathPara>", vbBinaryCompare)
        If mathParagraphEnd >= mathEnd Then
            ' A display wrapper may surround only this equation. Keep other
            ' math objects untouched rather than discarding a shared wrapper.
            If InStr(mathParagraphStart, sourceXml, "<m:oMath>", vbBinaryCompare) <> mathStart Or _
               InStr(mathEnd, sourceXml, "<m:oMath>", vbBinaryCompare) < mathParagraphEnd And _
               InStr(mathEnd, sourceXml, "<m:oMath>", vbBinaryCompare) > 0 Then GoTo InvalidCarrier
            mathStart = mathParagraphStart
            mathEnd = mathParagraphEnd + Len("</m:oMathPara>")
        End If
    End If
    ' Keep the original Flat OPC media/relationships, including SVG and its
    ' fallback. Only the selected math container becomes an ordinary drawing.
    sourceXml = Left$(sourceXml, mathStart - 1) & _
        "<w:r>" & drawingXml & "</w:r>" & Mid$(sourceXml, mathEnd)
    stagingPath = VTApplicationSupportRoot() & "/plain-image-" & VTNewUuidV4() & ".xml"
    VTWriteTextAtomic stagingPath, sourceXml
    Set parsedDocument = Documents.Open(FileName:=stagingPath, _
        AddToRecentFiles:=False, Visible:=False)
    If parsedDocument.InlineShapes.Count <> sourceParagraph.InlineShapes.Count Or _
       parsedDocument.OMaths.Count <> sourceParagraph.OMaths.Count - 1 Then GoTo InvalidCarrier
    If parsedDocument.InlineShapes(pictureOrdinal).Title <> originalTitle Or _
       parsedDocument.InlineShapes(pictureOrdinal).AlternativeText <> originalMetadata Or _
       parsedDocument.InlineShapes(pictureOrdinal).Range.OMaths.Count <> 0 Then GoTo InvalidCarrier
    Set replacementRange = parsedDocument.Paragraphs(1).Range.Duplicate

    Set backupDocument = Documents.Add(Visible:=False)
    Set backupRange = backupDocument.Content.Duplicate
    backupRange.FormattedText = sourceParagraph.FormattedText
    Set backupRange = backupDocument.Paragraphs(1).Range.Duplicate
    documentObject.Activate
    sourceReplaced = True
    targetDocumentMutated = True
    sourceParagraph.FormattedText = replacementRange.FormattedText
    replacementEnd = paragraphEnd + documentObject.Content.End - originalContentEnd
    Set replacementRange = documentObject.Range( _
        Start:=paragraphStart, End:=replacementEnd)
    If replacementRange.InlineShapes.Count <> pictureCount Then GoTo InvalidCarrier
    Set formulaShape = replacementRange.InlineShapes(pictureOrdinal)
    If formulaShape.Title <> originalTitle Or _
       formulaShape.AlternativeText <> originalMetadata Or _
       formulaShape.Range.OMaths.Count <> 0 Then GoTo InvalidCarrier
    VTDeleteTrailingNativeImageArtifact formulaShape
    parsedDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set parsedDocument = Nothing
    backupDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set backupDocument = Nothing
    On Error Resume Next
    Kill stagingPath
    On Error GoTo 0
    VTEndWordInternalMutation
    Set VTDetachWordFormulaPictureFromMath = formulaShape
    Exit Function

InvalidCarrier:
    Err.Raise vbObjectError + 7594, "VisualTeX", _
        "Word could not rebuild the converted image as an ordinary paragraph."
DetachFailed:
    detachErrorNumber = Err.Number
    detachErrorDescription = Err.Description
    On Error Resume Next
    If sourceReplaced And Not backupRange Is Nothing Then
        replacementEnd = paragraphEnd + documentObject.Content.End - originalContentEnd
        Set replacementRange = documentObject.Range( _
            Start:=paragraphStart, End:=replacementEnd)
        replacementRange.FormattedText = backupRange.FormattedText
    End If
    If Not parsedDocument Is Nothing Then parsedDocument.Close SaveChanges:=wdDoNotSaveChanges
    If Not backupDocument Is Nothing Then backupDocument.Close SaveChanges:=wdDoNotSaveChanges
    If Len(stagingPath) > 0 Then Kill stagingPath
    If internalMutationStarted Then VTEndWordInternalMutation
    On Error GoTo 0
    Err.Raise detachErrorNumber, "VisualTeX native-to-image cleanup", detachErrorDescription
End Function

Public Function VTFindUniqueInlineShape(ByVal marker As String) As InlineShape
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

Public Function VTFindCommittedInlineShapeNearPosition( _
    ByVal documentObject As Document, _
    ByVal expectedStart As Long, _
    ByVal metadata As String, _
    ByVal formulaReference As String, _
    Optional ByVal maximumDistance As Long = 64) As InlineShape

    Dim probeRange As Range
    Dim item As InlineShape
    Dim match As InlineShape
    Dim probeStart As Long
    Dim probeEnd As Long
    Dim count As Long

    If documentObject Is Nothing Or expectedStart < 0 Or _
       maximumDistance < 0 Then Exit Function
    probeStart = expectedStart - maximumDistance
    If probeStart < 0 Then probeStart = 0
    probeEnd = expectedStart + maximumDistance + 1
    If probeEnd > documentObject.Content.End Then
        probeEnd = documentObject.Content.End
    End If
    Set probeRange = documentObject.Range( _
        Start:=probeStart, End:=probeEnd)
    For Each item In probeRange.InlineShapes
        If item.AlternativeText = metadata And _
           item.Title = formulaReference Then
            count = count + 1
            Set match = item
        End If
    Next item
    If count > 1 Then
        Err.Raise vbObjectError + 7427, "VisualTeX", _
            "Word contains multiple local copies of the same VisualTeX Session result."
    End If
    If count = 1 Then
        Set VTFindCommittedInlineShapeNearPosition = match
    End If
End Function

Public Function VTFindCommittedInlineShapeInDocument( _
    ByVal documentObject As Document, _
    ByVal metadata As String, _
    ByVal formulaReference As String) As InlineShape

    Dim item As InlineShape
    Dim match As InlineShape
    Dim count As Long

    If documentObject Is Nothing Then Exit Function
    For Each item In documentObject.InlineShapes
        If item.AlternativeText = metadata And item.Title = formulaReference Then
            count = count + 1
            Set match = item
        End If
    Next item
    If count > 1 Then
        Err.Raise vbObjectError + 7427, "VisualTeX", _
            "Word contains multiple committed copies of the same VisualTeX Session result."
    End If
    If count = 1 Then
        Set VTFindCommittedInlineShapeInDocument = match
    End If
End Function

Public Function VTFindCommittedInlineShape( _
    ByVal metadata As String, _
    ByVal formulaReference As String) As InlineShape

    Set VTFindCommittedInlineShape = _
        VTFindCommittedInlineShapeInDocument( _
            ActiveDocument, metadata, formulaReference)
End Function

Public Function VTResolveImageFormulaByIdentity( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal metadata As String, _
    ByVal formulaReference As String, _
    Optional ByVal expectedStart As Long = -1) As InlineShape

    Dim formulaShape As InlineShape

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Or _
       Len(metadata) = 0 Or Len(formulaReference) = 0 Then
        Err.Raise vbObjectError + 7538, "VisualTeX", _
            "The numbered image identity is incomplete."
    End If

    ' VT_I_ is the durable image owner after a completed commit. During a fresh
    ' create transaction it may not exist yet, so use the last known local
    ' position and finally the exact metadata/reference pair. Every branch
    ' returns a newly fetched InlineShape after Word changed fields or paragraphs.
    Set formulaShape = VTWordImageOwnerShape(documentObject, formulaId)
    If Not formulaShape Is Nothing Then
        If formulaShape.AlternativeText = metadata And _
           formulaShape.Title = formulaReference Then
            Set VTResolveImageFormulaByIdentity = formulaShape
            Exit Function
        End If
        Set formulaShape = Nothing
    End If
    If expectedStart >= 0 Then
        Set formulaShape = VTFindCommittedInlineShapeNearPosition( _
            documentObject, expectedStart, metadata, formulaReference, 256)
    End If
    If formulaShape Is Nothing Then
        Set formulaShape = VTFindCommittedInlineShapeInDocument( _
            documentObject, metadata, formulaReference)
    End If
    If formulaShape Is Nothing Then
        Err.Raise vbObjectError + 7538, "VisualTeX", _
            "Word could not re-resolve exactly one numbered formula image."
    End If
    Set VTResolveImageFormulaByIdentity = formulaShape
End Function

Public Function VTExpectedImageFormulaPosition( _
    ByVal formulaShape As InlineShape, _
    ByVal baselineTextSize As Double, _
    ByVal valueLabel As String, _
    Optional ByVal numberFontName As String = "") As Long

    Dim formulaHeight As Double
    Dim inkCenterYRatio As Double
    Dim formulaInkCenterAboveBaseline As Double
    Dim baselineTextInkCenterAboveBaseline As Double
    Dim rawPosition As Double
    Dim inkTopEm As Double
    Dim inkBottomEm As Double
    Dim fontTopEm As Double
    Dim fontDescentEm As Double

    If formulaShape Is Nothing Then
        Err.Raise vbObjectError + 7564, "VisualTeX", _
            valueLabel & " target is missing."
    End If
    formulaHeight = CDbl(formulaShape.Height)
    If formulaHeight <= 0# Or formulaHeight > 10000# Then
        Err.Raise vbObjectError + 7564, "VisualTeX", _
            "The numbered image Equation has an invalid rendered height."
    End If
    If baselineTextSize <= 0# Or _
       baselineTextSize > VT_WORD_MAX_FORMULA_FONT_SIZE_PT Then
        baselineTextSize = CDbl(VTVisibleEquationNumberFontSize( _
            formulaShape.Range.Document))
    End If
    If baselineTextSize <= 0# Or _
       baselineTextSize > VT_WORD_MAX_FORMULA_FONT_SIZE_PT Then
        Err.Raise vbObjectError + 7564, "VisualTeX", _
            "The image Equation has an invalid baseline text size."
    End If

    inkCenterYRatio = VTNumberedImageInkCenterYRatio(formulaShape)
    formulaInkCenterAboveBaseline = formulaHeight * (1# - inkCenterYRatio)
    VTEquationNumberFontMetrics numberFontName, _
        inkTopEm, inkBottomEm, fontTopEm, fontDescentEm
    VTExpectedImageFormulaPosition = VTWordPicturePosition( _
        formulaShape, formulaInkCenterAboveBaseline, baselineTextSize, _
        inkTopEm, inkBottomEm, fontTopEm, fontDescentEm)
End Function

Public Sub VTApplyUnnumberedImageFormulaVerticalAlignment( _
    ByVal formulaRange As Range)

    Dim formulaShape As InlineShape
    Dim paragraphRange As Range
    Dim paragraphMarkRange As Range
    Dim baselineTextSize As Double
    Dim formulaPosition As Long

    If formulaRange Is Nothing Then Exit Sub
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    If paragraphRange.InlineShapes.Count <> 1 Or _
       paragraphRange.OMaths.Count <> 0 Then Exit Sub

    Set formulaShape = paragraphRange.InlineShapes(1)
    baselineTextSize = CDbl(VTVisibleEquationNumberFontSize( _
        paragraphRange.Document))
    formulaPosition = VTExpectedImageFormulaPosition( _
        formulaShape, baselineTextSize, _
        "The unnumbered image Equation paragraph-mark alignment")
    VTApplyWordImagePosition formulaShape, formulaPosition

    ' Keep Return/typing on Word's ordinary paragraph baseline. Moving only the
    ' picture gives an unnumbered display the same visual-centre relationship as
    ' the centre/right-tab numbered layout without contaminating the next line.
    If paragraphRange.End > paragraphRange.Start Then
        Set paragraphMarkRange = paragraphRange.Document.Range( _
            Start:=paragraphRange.End - 1, End:=paragraphRange.End)
        With paragraphMarkRange.Font
            If .Superscript <> False Then .Superscript = False
            If .Subscript <> False Then .Subscript = False
            If .Position <> 0 Then .Position = 0
        End With
    End If
End Sub

Public Sub VTNormalizeImageDisplayParagraph(ByVal formulaRange As Range)
    Dim paragraphRange As Range

    If formulaRange Is Nothing Then Exit Sub
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    ' Image display formulas own their visual height already. Extra paragraph
    ' spacing compounds that height and creates a much larger gap than native
    ' OMML. Keep this image-only so existing OMML paragraph spacing is untouched.
    With paragraphRange.ParagraphFormat
        .SpaceBefore = 0!
        .SpaceAfter = 0!
    End With
End Sub

Public Function VTImageCacheDocumentPath(ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "The cached Word image formula id is invalid."
    End If
    VTImageCacheDocumentPath = _
        VTApplicationSupportRoot() & "/ImageDocuments/" & formulaId & ".docx"
End Function

Public Function VTImageCachePngPath(ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "The cached Word image formula id is invalid."
    End If
    VTImageCachePngPath = _
        VTApplicationSupportRoot() & "/ImageDocuments/" & formulaId & ".png"
End Function

' Gives every formula of one kind ("inline" or "block") in the active
' document the chosen size, VisualTeX image formulas and Word equations alike,
' as one Undo step. A numbered equation is a display formula. Returns how many
' formulas were resized.
Public Function VTApplyFontSizeToDocumentFormulas( _
    ByVal wantedDisplayMode As String, _
    ByVal fontSizePt As Double) As Long

    Dim documentObject As Document
    Dim imageFormulas As New Collection
    Dim candidate As InlineShape
    Dim imageFormula As Variant
    Dim formulaMath As OMath
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim undoRecord As Object
    Dim undoRecordStarted As Boolean
    Dim internalMutationStarted As Boolean
    Dim previousScreenUpdating As Boolean
    Dim resized As Long
    Dim errorNumber As Long
    Dim errorDescription As String

    If Not VTValidWordFormulaFontSize(fontSizePt) Then Exit Function
    If Documents.Count = 0 Then Exit Function
    Set documentObject = ActiveDocument
    For Each candidate In documentObject.InlineShapes
        If VTTryWordImageTypographyReference( _
           candidate, formulaId, displayMode, numbered) Then
            If displayMode = wantedDisplayMode Then imageFormulas.Add candidate
        End If
    Next candidate

    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
    undoRecord.StartCustomRecord "VisualTeX Formula Font Size"
    undoRecordStarted = True
    VTBeginWordInternalMutation
    internalMutationStarted = True

    If imageFormulas.Count > 0 Then
        ' One request for the letter fonts their numbers take.
        VTEnsureWordImageLetterFonts documentObject, imageFormulas
        For Each imageFormula In imageFormulas
            VTApplyWordImageFormulaFontSize imageFormula, fontSizePt
            resized = resized + 1
        Next imageFormula
    End If

    For Each formulaMath In documentObject.OMaths
        If formulaMath.Type = wdOMathDisplay Then
            displayMode = "block"
        Else
            displayMode = "inline"
        End If
        If Not VTPureNativeNumberField(formulaMath.Range) Is Nothing Then _
            displayMode = "block"
        If displayMode = wantedDisplayMode Then
            VTV2ApplyEquationTypography formulaMath.Range, fontSizePt
            VTRememberNativeFormulaFontSize documentObject, formulaMath, fontSizePt
            resized = resized + 1
        End If
    Next formulaMath

    VTEndWordInternalMutation
    internalMutationStarted = False
    undoRecord.EndCustomRecord
    undoRecordStarted = False
    Application.ScreenUpdating = previousScreenUpdating
    VTApplyFontSizeToDocumentFormulas = resized
    Exit Function

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If internalMutationStarted Then VTEndWordInternalMutation
    If undoRecordStarted Then undoRecord.EndCustomRecord
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    Err.Raise errorNumber, "VisualTeX formula font size", errorDescription
End Function

' A VisualTeX equation keeps the size its image is drawn at for a later
' conversion back to an image: store the new size there too.
Private Sub VTRememberNativeFormulaFontSize( _
    ByVal documentObject As Document, _
    ByVal formulaMath As OMath, _
    ByVal fontSizePt As Double)

    Dim nativeBookmark As Bookmark
    Dim formulaId As String
    Dim storedFontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim observedWordFontSizePt As Double

    On Error GoTo Finished
    If Not VTTryFindNativeFormulaBookmarkLocally( _
       formulaMath.Range.Duplicate, nativeBookmark) Then Exit Sub
    If Not VTTryFormulaIdFromNativeBookmark( _
       nativeBookmark.Name, formulaId) Then Exit Sub
    If Not VTTryReadWordImageScaleState( _
       documentObject, formulaId, storedFontSizePt, referenceWidthPt, _
       referenceHeightPt, referenceBaselinePt, observedWordFontSizePt) Then Exit Sub
    VTSetWordImageScaleState _
        documentObject, formulaId, fontSizePt, referenceWidthPt, _
        referenceHeightPt, referenceBaselinePt, fontSizePt

Finished:
    Err.Clear
End Sub

Public Function VTWordImageScaleVariableName( _
    ByVal formulaId As String) As String

    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7490, "VisualTeX", _
            "VisualTeX cannot address image scale metadata for an invalid formula id."
    End If
    VTWordImageScaleVariableName = _
        VT_WORD_IMAGE_SCALE_VARIABLE_PREFIX & Replace$(formulaId, "-", "_")
End Function

Public Function VTWordImageCanvasVariableName(ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(formulaId) Then Err.Raise 5, "VisualTeX", "Invalid image canvas identity."
    VTWordImageCanvasVariableName = VT_WORD_IMAGE_CANVAS_VARIABLE_PREFIX & Replace$(formulaId, "-", "_")
End Function

Public Function VTWordImageGlyphHeight(ByVal formulaShape As InlineShape) As Double
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim savedValue As String
    Dim dimensions() As String
    VTWordImageGlyphHeight = formulaShape.Height
    If Not VTTryParseFormulaReference(formulaShape.Title, formulaId, displayMode, numbered) Then Exit Function
    If Not VTTryGetDocumentVariable(formulaShape.Range.Document, _
        VTWordImageCanvasVariableName(formulaId), savedValue) Then Exit Function
    If VTCachedWordImageCropTop(formulaShape) >= 0# Then Exit Function
    dimensions = Split(savedValue, "|")
    If UBound(dimensions) < 1 Or UBound(dimensions) > 2 Then Exit Function
    If Val(dimensions(0)) <= 0# Or Val(dimensions(1)) <= 0# Then Exit Function
    ' Negative CropTop adds transparent canvas; it does not resize painted SVG
    ' glyphs. Track frame height independently of width for native resizing.
    If UBound(dimensions) = 2 Then
        If Val(dimensions(2)) > 0# Then _
            VTWordImageGlyphHeight = Val(dimensions(1)) * formulaShape.Height / Val(dimensions(2))
    Else
        VTWordImageGlyphHeight = Val(dimensions(1)) * formulaShape.Width / Val(dimensions(0))
    End If
End Function

Public Function VTCachedWordImageCropTop(ByVal formulaShape As InlineShape) As Double
    Dim cacheKey As String
    Dim cachedValue As String
    Dim cropTop As Double
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim canvasState As String
    ' Merely reading PictureFormat.CropTop dirties a document in Word for Mac.
    ' Read it once per document/formula/frame geometry, then keep steady idle
    ' checks away from PictureFormat. Cropping/resizing changes the frame size.
    cacheKey = formulaShape.Range.Document.FullName & "|" & formulaShape.Title & _
        "|" & CStr(formulaShape.Width) & ":" & CStr(formulaShape.Height)
    ' Returning to an old frame size after removing padding must not reuse the
    ' old negative crop. Canvas metadata distinguishes those frame histories.
    If VTTryParseFormulaReference(formulaShape.Title, formulaId, displayMode, numbered) Then
        VTTryGetDocumentVariable formulaShape.Range.Document, _
            VTWordImageCanvasVariableName(formulaId), canvasState
    End If
    cacheKey = cacheKey & "|" & canvasState
    If VT_WORD_IMAGE_CROP_CACHE Is Nothing Then _
        Set VT_WORD_IMAGE_CROP_CACHE = New Collection
    On Error Resume Next
    cachedValue = VT_WORD_IMAGE_CROP_CACHE(cacheKey)
    Err.Clear
    On Error GoTo 0
    If Len(cachedValue) > 0 Then
        VTCachedWordImageCropTop = CDbl(cachedValue)
        Exit Function
    End If
    cropTop = formulaShape.PictureFormat.CropTop
    If VT_WORD_IMAGE_CROP_CACHE.Count >= 512 Then _
        Set VT_WORD_IMAGE_CROP_CACHE = New Collection
    VT_WORD_IMAGE_CROP_CACHE.Add CStr(cropTop), cacheKey
    VTCachedWordImageCropTop = cropTop
End Function

Public Sub VTResetWordImageAlignmentCanvas(ByVal formulaShape As InlineShape)
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim savedValue As String
    Dim glyphHeight As Double
    Dim glyphWidth As Double
    Dim oldAspectLock As Long
    If Not VTTryParseFormulaReference(formulaShape.Title, formulaId, displayMode, numbered) Then Exit Sub
    If Not VTTryGetDocumentVariable(formulaShape.Range.Document, _
        VTWordImageCanvasVariableName(formulaId), savedValue) Then Exit Sub
    If VTCachedWordImageCropTop(formulaShape) >= 0# Then Exit Sub
    ' Only reset the transparent canvas owned by this solver. Word rounds
    ' uncropping separately; restore the recorded glyph dimensions to prevent
    ' accumulating a quarter-point of shrinkage on repeated mode changes.
    If VTCachedWordImageCropTop(formulaShape) < 0# Then
        glyphHeight = VTWordImageGlyphHeight(formulaShape)
        glyphWidth = formulaShape.Width
        oldAspectLock = formulaShape.LockAspectRatio
        formulaShape.LockAspectRatio = msoFalse
        formulaShape.PictureFormat.CropTop = 0!
        formulaShape.Width = CSng(glyphWidth)
        formulaShape.Height = CSng(glyphHeight)
        formulaShape.LockAspectRatio = oldAspectLock
    End If
    VTDeleteDocumentVariable formulaShape.Range.Document, VTWordImageCanvasVariableName(formulaId)
End Sub

Public Function VTWordImageInkCenterVariableName( _
    ByVal formulaId As String) As String

    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7490, "VisualTeX", _
            "VisualTeX cannot address image ink geometry for an invalid formula id."
    End If
    VTWordImageInkCenterVariableName = _
        VT_WORD_IMAGE_INK_CENTER_VARIABLE_PREFIX & _
        Replace$(formulaId, "-", "_")
End Function

Public Sub VTSetWordImageInkCenterYRatio( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal inkCenterYRatio As Double)

    If documentObject Is Nothing Or _
       inkCenterYRatio < 0# Or inkCenterYRatio > 1# Then
        Err.Raise vbObjectError + 7491, "VisualTeX", _
            "VisualTeX image ink-center metadata is invalid."
    End If
    VTSetDocumentVariable _
        documentObject, _
        VTWordImageInkCenterVariableName(formulaId), _
        VTJsonNumber(inkCenterYRatio)
End Sub

Public Function VTTryReadWordImageInkCenterYRatio( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByRef inkCenterYRatio As Double) As Boolean

    Dim storedValue As String

    inkCenterYRatio = 0.5
    If documentObject Is Nothing Then Exit Function
    If Not VTTryGetDocumentVariable( _
        documentObject, VTWordImageInkCenterVariableName(formulaId), _
        storedValue) Then Exit Function
    inkCenterYRatio = Val(storedValue)
    If inkCenterYRatio < 0# Or inkCenterYRatio > 1# Then
        inkCenterYRatio = 0.5
        Exit Function
    End If
    VTTryReadWordImageInkCenterYRatio = True
End Function

Public Function VTWordImageLetterFontVariableName( _
    ByVal formulaId As String) As String

    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7490, "VisualTeX", _
            "VisualTeX cannot address an image letter font for an invalid formula id."
    End If
    VTWordImageLetterFontVariableName = _
        VT_WORD_IMAGE_LETTER_FONT_VARIABLE_PREFIX & _
        Replace$(formulaId, "-", "_")
End Function

Private Function VTIsKnownFormulaLetterFont(ByVal letterFont As String) As Boolean
    Select Case letterFont
        Case "katex", "times", "cambria", "stix", "palatino", "helvetica"
            VTIsKnownFormulaLetterFont = True
    End Select
End Function

' The formula's letter font as VisualTeX names it (katex, times, cambria,
' stix, palatino, helvetica), kept so its Equation number can use it too.
Public Sub VTSetWordImageLetterFont( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal letterFont As String)

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then Exit Sub
    If Not VTIsKnownFormulaLetterFont(letterFont) Then Exit Sub
    VTSetDocumentVariable documentObject, _
        VTWordImageLetterFontVariableName(formulaId), letterFont
End Sub

Private Function VTHasWordImageLetterFont( _
    ByVal documentObject As Document, _
    ByVal formulaId As String) As Boolean

    Dim letterFont As String
    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then Exit Function
    VTHasWordImageLetterFont = VTTryGetDocumentVariable(documentObject, _
        VTWordImageLetterFontVariableName(formulaId), letterFont)
End Function

' Records the letter font of every formula in formulaShapes that has none yet,
' with one request to the app, which reads it from the formula's metadata.
' Formulas from before letter fonts were recorded have none. A formula the app
' cannot read keeps none (its number stays in Cambria Math).
Private Sub VTEnsureWordImageLetterFonts( _
    ByVal documentObject As Document, _
    ByVal formulaShapes As Collection)

    Dim shape As Variant
    Dim formulaShape As InlineShape
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim metadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean
    Dim formulaIds As New Collection
    Dim manifest As String
    Dim sessionId As String
    Dim timingDetail As String
    Dim separatorPosition As Long
    Dim result As Object
    Dim index As Long
    Dim key As String

    On Error GoTo LetterFontsUnavailable
    If documentObject Is Nothing Or formulaShapes Is Nothing Then Exit Sub
    For Each shape In formulaShapes
        Set formulaShape = shape
        ' The metadata as batch numbering reads it (never repairing the shape).
        If VTTryResolveVisualTeXInlineShapeReference( _
           formulaShape, formulaId, displayMode, numbered, metadata, _
           metadataNeedsWrite, formatNeedsWrite, False) Then
            If Not VTHasWordImageLetterFont(documentObject, formulaId) And _
               VTIsEncodedMetadata(metadata) Then
                formulaIds.Add formulaId
                manifest = manifest & "image" & CStr(formulaIds.Count) & _
                    "Metadata=" & metadata & vbLf
            End If
        End If
    Next shape
    If formulaIds.Count = 0 Then Exit Sub

    sessionId = VTNewUuidV4()
    manifest = "protocolVersion=" & CStr(VT_PROTOCOL_VERSION) & vbLf & _
        "sessionId=" & sessionId & vbLf & _
        "sourceKind=imageLetterFonts" & vbLf & _
        "sourceDocumentId=" & VTBase64UrlEncodeUtf8(documentObject.FullName) & vbLf & _
        "bookmarkName=-" & vbLf & _
        "imageCount=" & CStr(formulaIds.Count) & vbLf & manifest
    If VTTryWriteOmmlBatchManifestDirect(sessionId, manifest) Then
        timingDetail = VTConvertOmmlBatch(VT_WORD_HOST, sessionId)
    Else
        timingDetail = VTConvertOmmlBatch(VT_WORD_HOST, sessionId, manifest)
    End If
    separatorPosition = InStrRev(timingDetail, "|", -1, vbBinaryCompare)
    If separatorPosition <= 1 Or separatorPosition >= Len(timingDetail) Then Exit Sub
    Set result = VTParseDocumentImportManifestText( _
        VTBase64UrlDecodeUtf8(Mid$(timingDetail, separatorPosition + 1)), sessionId)
    If CStr(result("operation")) <> "imageLetterFonts" Then Exit Sub
    For index = 1 To formulaIds.Count
        key = "image" & CStr(index) & "LetterFont"
        If VTCollectionHasKey(result, key) Then
            VTSetWordImageLetterFont documentObject, formulaIds(index), CStr(result(key))
        Else
            ' Its metadata names no letter font: remember that, so the app is
            ' not asked again; the number stays in Cambria Math.
            VTSetDocumentVariable documentObject, _
                VTWordImageLetterFontVariableName(CStr(formulaIds(index))), "-"
        End If
    Next index
    Exit Sub

LetterFontsUnavailable:
    ' The number keeps Cambria Math; nothing else depends on the letter font.
    Err.Clear
End Sub

Public Sub VTEnsureWordImageLetterFont(ByVal formulaShape As InlineShape)
    Dim shapes As New Collection
    If formulaShape Is Nothing Then Exit Sub
    shapes.Add formulaShape
    VTEnsureWordImageLetterFonts formulaShape.Range.Document, shapes
End Sub

Public Function VTTryReadCachedWordImageInkCenterYRatio( _
    ByVal formulaId As String, _
    ByRef inkCenterYRatio As Double) As Boolean

    Dim response As String
    Dim numericValue As Double

    inkCenterYRatio = 0.5
    If Not VTIsCanonicalUuid(formulaId) Then Exit Function
    On Error GoTo CacheUnavailable
    response = VTFileBridgeCall( _
        "ReadVisualTeXImageInkCenter", formulaId)
    numericValue = Val(Trim$(response))
    If numericValue < 0# Or numericValue > 1# Then Exit Function
    inkCenterYRatio = numericValue
    VTTryReadCachedWordImageInkCenterYRatio = True
    Exit Function

CacheUnavailable:
    Err.Clear
    inkCenterYRatio = 0.5
End Function

Public Function VTValidWordFormulaFontSize(ByVal value As Double) As Boolean
    VTValidWordFormulaFontSize = _
        value >= VT_WORD_MIN_FORMULA_FONT_SIZE_PT And _
        value <= VT_WORD_MAX_FORMULA_FONT_SIZE_PT
End Function

Public Function VTPreferredWordFormulaFontSize( _
    ByVal contextRange As Range) As Double

    Dim contextSize As Double
    Dim normalSize As Double

    On Error Resume Next
    If Not contextRange Is Nothing Then contextSize = contextRange.Font.Size
    If Not contextRange Is Nothing Then
        normalSize = contextRange.Document.Styles(wdStyleNormal).Font.Size
    ElseIf Documents.Count > 0 Then
        normalSize = ActiveDocument.Styles(wdStyleNormal).Font.Size
    End If
    On Error GoTo 0

    If VTValidWordFormulaFontSize(contextSize) Then
        VTPreferredWordFormulaFontSize = contextSize
    ElseIf VTValidWordFormulaFontSize(normalSize) Then
        VTPreferredWordFormulaFontSize = normalSize
    Else
        VTPreferredWordFormulaFontSize = VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
    End If
End Function

Public Sub VTSetWordImageScaleState( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal fontSizePt As Double, _
    ByVal referenceWidthPt As Double, _
    ByVal referenceHeightPt As Double, _
    ByVal referenceBaselinePt As Double, _
    ByVal observedWordFontSizePt As Double)

    If documentObject Is Nothing Or _
       Not VTValidWordFormulaFontSize(fontSizePt) Or _
       referenceWidthPt <= 0# Or referenceHeightPt <= 0# Or _
       referenceBaselinePt < -256# Or referenceBaselinePt > 0# Then
        Err.Raise vbObjectError + 7491, "VisualTeX", _
            "VisualTeX image scale metadata is invalid."
    End If
    If Not VTValidWordFormulaFontSize(observedWordFontSizePt) Then
        observedWordFontSizePt = fontSizePt
    End If

    VTSetDocumentVariable _
        documentObject, _
        VTWordImageScaleVariableName(formulaId), _
        VTJsonNumber(fontSizePt) & "|" & _
        VTJsonNumber(referenceWidthPt) & "|" & _
        VTJsonNumber(referenceHeightPt) & "|" & _
        VTJsonNumber(referenceBaselinePt) & "|" & _
        VTJsonNumber(observedWordFontSizePt)
End Sub

Public Function VTWordImageScaleStateUsesPlaceholderGeometry( _
    ByVal referenceWidthPt As Double, _
    ByVal referenceHeightPt As Double, _
    ByVal referenceBaselinePt As Double) As Boolean

    VTWordImageScaleStateUsesPlaceholderGeometry = _
        Abs(referenceWidthPt - VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT) <= 0.01 And _
        Abs(referenceHeightPt - VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT) <= 0.01 And _
        Abs(referenceBaselinePt) <= 0.01
End Function

Public Function VTTryReadWordImageScaleState( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByRef fontSizePt As Double, _
    ByRef referenceWidthPt As Double, _
    ByRef referenceHeightPt As Double, _
    ByRef referenceBaselinePt As Double, _
    ByRef observedWordFontSizePt As Double) As Boolean

    Dim storedValue As String
    Dim fields() As String

    fontSizePt = 0#
    referenceWidthPt = 0#
    referenceHeightPt = 0#
    referenceBaselinePt = 0#
    observedWordFontSizePt = 0#
    If Not VTTryGetDocumentVariable( _
        documentObject, VTWordImageScaleVariableName(formulaId), storedValue) Then Exit Function
    fields = Split(storedValue, "|")
    If UBound(fields) <> 4 Then GoTo InvalidScale
    fontSizePt = Val(fields(0))
    referenceWidthPt = Val(fields(1))
    referenceHeightPt = Val(fields(2))
    referenceBaselinePt = Val(fields(3))
    observedWordFontSizePt = Val(fields(4))
    If Not VTValidWordFormulaFontSize(fontSizePt) Or _
       referenceWidthPt <= 0# Or referenceHeightPt <= 0# Or _
       referenceBaselinePt < -256# Or referenceBaselinePt > 0# Or _
       Not VTValidWordFormulaFontSize(observedWordFontSizePt) Then _
        GoTo InvalidScale
    VTTryReadWordImageScaleState = True
    Exit Function

InvalidScale:
    Err.Raise vbObjectError + 7492, "VisualTeX", _
        "The stored VisualTeX image scale metadata is invalid or corrupt."
End Function

Public Function VTInlineShapeWordFontSize( _
    ByVal formulaShape As InlineShape) As Double

    Dim value As Double

    On Error Resume Next
    value = formulaShape.Range.Font.Size
    On Error GoTo 0
    If VTValidWordFormulaFontSize(value) Then
        VTInlineShapeWordFontSize = value
    Else
        VTInlineShapeWordFontSize = 0#
    End If
End Function

Public Function VTInlineShapeFontPosition( _
    ByVal formulaShape As InlineShape) As Double

    Dim value As Double

    On Error Resume Next
    value = formulaShape.Range.Font.Position
    On Error GoTo 0
    If value < -256# Then value = -256#
    If value > 0# Then value = 0#
    VTInlineShapeFontPosition = value
End Function

Public Sub VTEnsureWordImageScaleState( _
    ByVal formulaShape As InlineShape, _
    ByRef formulaId As String, _
    ByRef displayMode As String, _
    ByRef numbered As Boolean, _
    ByRef fontSizePt As Double, _
    ByRef referenceWidthPt As Double, _
    ByRef referenceHeightPt As Double, _
    ByRef referenceBaselinePt As Double, _
    ByRef observedWordFontSizePt As Double)

    Dim actualWordFontSize As Double
    Dim currentBaseline As Double
    Dim expectedWidth As Double
    Dim expectedHeight As Double
    Dim reportedWidth As Double
    Dim reportedHeight As Double
    Dim previousReferenceHeightPt As Double
    Dim baselineRatio As Double
    Dim adoptGeometry As Boolean

    If formulaShape Is Nothing Or _
       Not VTTryParseFormulaReference( _
           formulaShape.Title, formulaId, displayMode, numbered) Then
        Err.Raise vbObjectError + 7493, "VisualTeX", _
            "The selected VisualTeX image formula identity is invalid."
    End If
    actualWordFontSize = VTInlineShapeWordFontSize(formulaShape)
    currentBaseline = VTInlineShapeFontPosition(formulaShape)

    If VTTryReadWordImageScaleState( _
       formulaShape.Range.Document, formulaId, fontSizePt, _
       referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
       observedWordFontSizePt) Then
        expectedWidth = referenceWidthPt * _
            fontSizePt / VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
        expectedHeight = referenceHeightPt * _
            fontSizePt / VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
        If Abs(formulaShape.Width - expectedWidth) > 0.5 Or _
           Abs(VTWordImageGlyphHeight(formulaShape) - expectedHeight) > 0.5 Then
            adoptGeometry = True
            ' Some Word builds report the surrounding paragraph's Font.Size
            ' while an InlineShape is selected. Treat a changed Font.Size as a
            ' real native Home-size operation only when the picture geometry has
            ' also moved to the dimensions predicted by that size. Any other
            ' geometry change is authoritative image resizing and must not be
            ' rescaled again from a transient font report.
            If VTValidWordFormulaFontSize(actualWordFontSize) And _
               Abs(actualWordFontSize - observedWordFontSizePt) > 0.05 Then
                reportedWidth = referenceWidthPt * _
                    actualWordFontSize / VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
                reportedHeight = referenceHeightPt * _
                    actualWordFontSize / VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
                If Abs(formulaShape.Width - reportedWidth) <= 0.5 And _
                   Abs(VTWordImageGlyphHeight(formulaShape) - reportedHeight) <= 0.5 Then
                    adoptGeometry = False
                End If
            End If
            If adoptGeometry Then
                previousReferenceHeightPt = referenceHeightPt
                If previousReferenceHeightPt > 0# Then
                    baselineRatio = _
                        referenceBaselinePt / previousReferenceHeightPt
                Else
                    baselineRatio = 0#
                End If
                referenceWidthPt = formulaShape.Width * _
                    VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT / fontSizePt
                referenceHeightPt = VTWordImageGlyphHeight(formulaShape) * _
                    VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT / fontSizePt
                If baselineRatio < 0# Then
                    ' Preserve the SVG mathematical-baseline ratio for both
                    ' inline and block images. The previous inline path rebuilt
                    ' this value from an already-rounded Font.Position after a
                    ' resize, permanently introducing up to a point of drift.
                    referenceBaselinePt = _
                        referenceHeightPt * baselineRatio
                Else
                    referenceBaselinePt = 0#
                End If
                If referenceBaselinePt < -256# Then referenceBaselinePt = -256#
                If referenceBaselinePt > 0# Then referenceBaselinePt = 0#
                VTSetWordImageScaleState _
                    formulaShape.Range.Document, formulaId, fontSizePt, _
                    referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
                    observedWordFontSizePt
                If displayMode = "block" And numbered Then
                    VTRefreshNumberedImageFormulaAfterGeometryChange _
                        formulaShape, formulaId, fontSizePt
                End If
            End If
        End If
        Exit Sub
    End If

    If Not VTValidWordFormulaFontSize(actualWordFontSize) Then
        actualWordFontSize = _
            VTPreferredWordFormulaFontSize(formulaShape.Range)
    End If
    fontSizePt = actualWordFontSize
    observedWordFontSizePt = actualWordFontSize
    referenceWidthPt = formulaShape.Width * _
        VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT / fontSizePt
    referenceHeightPt = VTWordImageGlyphHeight(formulaShape) * _
        VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT / fontSizePt
    referenceBaselinePt = currentBaseline * _
        VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT / fontSizePt
    If referenceBaselinePt < -256# Then referenceBaselinePt = -256#
    If referenceBaselinePt > 0# Then referenceBaselinePt = 0#
    VTSetWordImageScaleState _
        formulaShape.Range.Document, formulaId, fontSizePt, _
        referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
        observedWordFontSizePt
    If displayMode = "block" And numbered Then
        VTRefreshNumberedImageFormulaAfterGeometryChange _
            formulaShape, formulaId, fontSizePt
    End If
End Sub

Private Function VTWordPicturePosition( _
    ByVal formulaShape As InlineShape, _
    ByVal anchorAboveBottomPt As Double, _
    ByVal textSizePt As Double, _
    ByVal inkTopEm As Double, _
    ByVal inkBottomEm As Double, _
    ByVal fontTopEm As Double, _
    ByVal fontDescentEm As Double) As Long

    ' Display (block) image formulas only: centre the formula's painted ink
    ' on the Equation number's painted centre. Inline formulas use
    ' VTWordInlineImagePosition.
    Dim rawPosition As Double
    Dim frameHeightPt As Double
    Dim textHalfInk As Double
    frameHeightPt = formulaShape.Height
    textHalfInk = (inkTopEm - inkBottomEm) * textSizePt / 2#
    Select Case formulaShape.Range.ParagraphFormat.BaseLineAlignment
        Case wdBaselineAlignTop
            rawPosition = frameHeightPt - anchorAboveBottomPt - textHalfInk
            If fontTopEm >= inkTopEm Then _
                rawPosition = rawPosition - (fontTopEm - inkTopEm) * textSizePt
        Case wdBaselineAlignCenter
            rawPosition = frameHeightPt / 2# - anchorAboveBottomPt
            If fontTopEm > 0# Then rawPosition = rawPosition + _
                (inkTopEm + inkBottomEm - fontTopEm + fontDescentEm) * textSizePt / 2#
        Case wdBaselineAlignFarEast50
            rawPosition = textHalfInk - anchorAboveBottomPt
            If fontTopEm > 0# Then rawPosition = _
                (fontTopEm - fontDescentEm) * textSizePt / 2# - anchorAboveBottomPt
        Case Else
            rawPosition = (inkTopEm + inkBottomEm) * textSizePt / 2# - anchorAboveBottomPt
    End Select
    If rawPosition < -256# Then rawPosition = -256#
    If rawPosition > 256# Then rawPosition = 256#
    ' Word quantizes picture height separately from Position. A 0.01 pt
    ' tie tolerance keeps adjacent pictures on the same baseline when their
    ' independently rounded heights straddle a mathematical half point.
    If rawPosition >= 0# Then
        VTWordPicturePosition = CLng(Int(rawPosition + 0.51#))
    Else
        VTWordPicturePosition = -CLng(Int(-rawPosition + 0.49#))
    End If
End Function

Public Function VTTryWordImageTypographyReference( _
    ByVal formulaShape As InlineShape, ByRef formulaId As String, _
    ByRef displayMode As String, ByRef numbered As Boolean, _
    Optional ByVal repairShapeProperties As Boolean = False) As Boolean
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean
    ' Reuse the normal image identity/recovery contract. SVG round-trips can
    ' remove or swap the compact title while retaining metadata and Variables.
    VTTryWordImageTypographyReference = VTTryResolveVisualTeXInlineShapeReference( _
        formulaShape, formulaId, displayMode, numbered, encodedMetadata, _
        metadataNeedsWrite, formatNeedsWrite, repairShapeProperties)
End Function

Private Function VTSemanticWordFormulaFontSize( _
    ByVal fontSizePt As Double) As Double

    ' Windows FormulaFontSize.Normalize: 14 pt when unknown, 5 to 200 pt,
    ' centipoint precision.
    If fontSizePt <= 0# Then fontSizePt = VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
    If fontSizePt < 5# Then fontSizePt = 5#
    If fontSizePt > 200# Then fontSizePt = 200#
    VTSemanticWordFormulaFontSize = Int(fontSizePt * 100# + 0.5) / 100#
End Function

' Inline image formulas follow the Windows rule
' (WindowsOffice.Contracts WordInlineAlignment.CalculateFontPositionForHost):
' - Auto and Baseline paragraphs: lower the picture by the formula's descent,
'   so the TeX baseline sits on the text baseline. The SVG pads the descent to
'   whole points, which is the only resolution Font.Position has.
' - Center and Top paragraphs: Word aligns the character boxes; Position 0.
' - FarEast50 (Bottom): lower by half of the picture height beyond the
'   formula's font size.
' The result depends only on the picture, its stored geometry and the
' paragraph mode, so reopening a document or editing nearby text never
' changes it.
Public Function VTWordInlineImagePosition( _
    ByVal formulaShape As InlineShape, _
    ByVal referenceHeightPt As Double, _
    ByVal referenceBaselinePt As Double, _
    ByVal formulaFontSizePt As Double) As Long

    Dim heightPt As Double
    Dim descentPt As Double
    Dim excessHeightPt As Double

    If formulaShape Is Nothing Or referenceHeightPt <= 0# Or _
       referenceBaselinePt < -256# Or referenceBaselinePt > 0# Then
        Err.Raise vbObjectError + 7494, "VisualTeX", _
            "The VisualTeX inline formula baseline reference is invalid."
    End If
    heightPt = VTWordImageGlyphHeight(formulaShape)
    If heightPt <= 0# Or heightPt > 10000# Then
        Err.Raise vbObjectError + 7494, "VisualTeX", _
            "The VisualTeX inline formula height is invalid."
    End If
    Select Case formulaShape.Range.ParagraphFormat.BaseLineAlignment
        Case wdBaselineAlignTop, wdBaselineAlignCenter
            VTWordInlineImagePosition = 0
        Case wdBaselineAlignFarEast50
            excessHeightPt = heightPt - _
                VTSemanticWordFormulaFontSize(formulaFontSizePt)
            If excessHeightPt < 0# Then excessHeightPt = 0#
            VTWordInlineImagePosition = -CLng(Int(excessHeightPt * 0.5 + 0.5))
        Case 2, 4 ' wdBaselineAlignBaseline, wdBaselineAlignAuto
            descentPt = -referenceBaselinePt * heightPt / referenceHeightPt
            If descentPt > 256# Then descentPt = 256#
            VTWordInlineImagePosition = -CLng(Int(descentPt + 0.5))
        Case Else
            Err.Raise vbObjectError + 7494, "VisualTeX", _
                "The inline formula's paragraph has no defined character alignment."
    End Select
End Function

Public Sub VTApplyWordInlineImageBaseline( _
    ByVal formulaShape As InlineShape, _
    ByVal referenceHeightPt As Double, _
    ByVal referenceBaselinePt As Double, _
    ByVal formulaFontSizePt As Double)

    Dim requestedPosition As Long
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim savedCanvas As String
    Dim ownsCanvas As Boolean
    Dim containingField As Field
    Dim restoreFieldLock As Boolean
    Dim mutationStarted As Boolean
    Dim applyErrorNumber As Long
    Dim applyErrorDescription As String
    On Error GoTo AlignmentFailed

    If formulaShape Is Nothing Then Exit Sub
    ' Releases before r117 gave formulas in Top paragraphs a transparent
    ' top canvas (a negative CropTop). Remove it once; the rule above aligns
    ' the plain picture.
    If VTTryParseFormulaReference(formulaShape.Title, formulaId, displayMode, numbered) Then
        ownsCanvas = VTTryGetDocumentVariable( _
            formulaShape.Range.Document, VTWordImageCanvasVariableName(formulaId), savedCanvas)
        If ownsCanvas Then ownsCanvas = VTCachedWordImageCropTop(formulaShape) < 0#
    End If
    requestedPosition = VTWordInlineImagePosition( _
        formulaShape, referenceHeightPt, referenceBaselinePt, formulaFontSizePt)
    If Not ownsCanvas And _
       formulaShape.Range.Font.Position = requestedPosition And _
       Not formulaShape.Range.Font.Superscript And _
       Not formulaShape.Range.Font.Subscript Then Exit Sub

    VTBeginWordInternalMutation
    mutationStarted = True
    If ownsCanvas Then
        Set containingField = VTVisualTeXImageMacroButtonFieldForShape(formulaShape)
        If Not containingField Is Nothing Then
            restoreFieldLock = containingField.Locked
            If restoreFieldLock Then containingField.Locked = False
        End If
        VTResetWordImageAlignmentCanvas formulaShape
        If Not containingField Is Nothing And restoreFieldLock Then containingField.Locked = True
        requestedPosition = VTWordInlineImagePosition( _
            formulaShape, referenceHeightPt, referenceBaselinePt, formulaFontSizePt)
    End If
    VTApplyWordImagePosition formulaShape, requestedPosition
    VTEndWordInternalMutation
    Exit Sub

AlignmentFailed:
    applyErrorNumber = Err.Number
    applyErrorDescription = Err.Description
    On Error Resume Next
    If Not containingField Is Nothing And restoreFieldLock Then containingField.Locked = True
    If mutationStarted Then VTEndWordInternalMutation
    On Error GoTo 0
    Err.Raise applyErrorNumber, "VisualTeX picture alignment", applyErrorDescription
End Sub

Public Sub VTApplyWordImagePosition( _
    ByVal formulaShape As InlineShape, _
    ByVal requestedPosition As Long)

    Dim appliedPosition As Long
    Dim attemptIndex As Long
    Dim writeErrorNumber As Long
    Dim applyErrorNumber As Long
    Dim applyErrorDescription As String
    Dim containingField As Field
    Dim restoreFieldLock As Boolean
    If formulaShape.Range.Font.Position = requestedPosition And _
       Not formulaShape.Range.Font.Superscript And _
       Not formulaShape.Range.Font.Subscript Then Exit Sub

    ' A committed image may live inside the locked MACROBUTTON result used for
    ' double-click editing. Word for Mac accepts the initial Position before the
    ' field is locked, but some builds ignore later Range.Font.Position writes to
    ' a locked result after an image resize. Temporarily unlock only our verified
    ' VisualTeX field, restore the lock on every exit path, and leave unrelated
    ' Word fields untouched.
    On Error GoTo ApplyFailed
    Set containingField = VTVisualTeXImageMacroButtonFieldForShape(formulaShape)
    If Not containingField Is Nothing Then
        restoreFieldLock = containingField.Locked
        If restoreFieldLock Then containingField.Locked = False
    End If

    ' An InlineShape can inherit raised/subscript formatting from the insertion
    ' caret. Clear those independent run flags before assigning Position so the
    ' same mathematical baseline is not offset a second time on another Word
    ' build or in a differently formatted paragraph.
    For attemptIndex = 1 To 3
        On Error Resume Next
        Err.Clear
        With formulaShape.Range.Font
            .Superscript = False
            .Subscript = False
            .Position = requestedPosition
        End With
        writeErrorNumber = Err.Number
        Err.Clear
        appliedPosition = formulaShape.Range.Font.Position
        If writeErrorNumber = 0 And Err.Number = 0 And _
           appliedPosition = requestedPosition Then
            On Error GoTo ApplyFailed
            If Not containingField Is Nothing And restoreFieldLock Then _
                containingField.Locked = True
            Exit Sub
        End If
        Err.Clear
        On Error GoTo ApplyFailed
        ' Word can keep the old Position while an InlineShape resize/layout is
        ' still being committed. Drain that host work before retrying the exact
        ' same mathematical baseline instead of accepting a stale value.
        DoEvents
    Next attemptIndex

    Err.Raise vbObjectError + 7494, "VisualTeX", _
        "Word did not persist the VisualTeX inline formula baseline" & _
        " [expected=" & CStr(requestedPosition) & _
        "; actual=" & CStr(appliedPosition) & "]."

ApplyFailed:
    applyErrorNumber = Err.Number
    applyErrorDescription = Err.Description
    On Error Resume Next
    If Not containingField Is Nothing And restoreFieldLock Then _
        containingField.Locked = True
    On Error GoTo 0
    If applyErrorNumber = 0 Then applyErrorNumber = vbObjectError + 7494
    If Len(applyErrorDescription) = 0 Then _
        applyErrorDescription = "Word could not update the VisualTeX inline formula baseline."
    Err.Raise applyErrorNumber, "VisualTeX inline baseline", _
        applyErrorDescription
End Sub

Public Sub VTApplyWordImageFormulaFontSize( _
    ByVal formulaShape As InlineShape, _
    ByVal requestedFontSizePt As Double)

    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim storedFontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim observedWordFontSizePt As Double
    Dim persistedWordFontSizePt As Double
    Dim targetWidth As Double
    Dim targetHeight As Double
    Dim targetBaseline As Double
    Dim internalMutationStarted As Boolean
    Dim applyErrorNumber As Long
    Dim applyErrorDescription As String

    If formulaShape Is Nothing Or _
       Not VTValidWordFormulaFontSize(requestedFontSizePt) Then
        Err.Raise vbObjectError + 7494, "VisualTeX", _
            "Enter a VisualTeX formula font size from 1 to 512 pt."
    End If
    VTEnsureWordImageScaleState _
        formulaShape, formulaId, displayMode, numbered, storedFontSizePt, _
        referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
        observedWordFontSizePt

    targetWidth = referenceWidthPt * _
        requestedFontSizePt / VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
    targetHeight = referenceHeightPt * _
        requestedFontSizePt / VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
    targetBaseline = referenceBaselinePt * _
        requestedFontSizePt / VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
    If targetBaseline < -256# Then targetBaseline = -256#
    If targetBaseline > 0# Then targetBaseline = 0#
    If targetWidth <= 0# Or targetHeight <= 0# Or _
       targetWidth > 10000# Or targetHeight > 10000# Then
        Err.Raise vbObjectError + 7495, "VisualTeX", _
            "The requested formula font size produces unsupported image dimensions."
    End If

    On Error GoTo ApplyFailed
    VTBeginWordInternalMutation
    internalMutationStarted = True
    VTResetWordImageAlignmentCanvas formulaShape
    On Error Resume Next
    formulaShape.Range.Font.Size = CSng(requestedFontSizePt)
    Err.Clear
    On Error GoTo ApplyFailed
    formulaShape.LockAspectRatio = msoFalse
    formulaShape.Width = CSng(targetWidth)
    formulaShape.Height = CSng(targetHeight)
    formulaShape.LockAspectRatio = msoTrue
    If displayMode = "inline" Then
        VTApplyWordInlineImageBaseline _
            formulaShape, referenceHeightPt, referenceBaselinePt, requestedFontSizePt
    Else
        formulaShape.Range.Font.Position = 0
        If numbered Then
            VTRefreshNumberedImageFormulaFontLayout _
                formulaShape, formulaId, requestedFontSizePt
        Else
            VTApplyUnnumberedImageFormulaVerticalAlignment _
                formulaShape.Range
        End If
    End If
    persistedWordFontSizePt = VTInlineShapeWordFontSize(formulaShape)
    If Not VTValidWordFormulaFontSize(persistedWordFontSizePt) Then
        persistedWordFontSizePt = observedWordFontSizePt
    End If
    VTSetWordImageScaleState _
        formulaShape.Range.Document, formulaId, requestedFontSizePt, _
        referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
        persistedWordFontSizePt
    VTEndWordInternalMutation
    internalMutationStarted = False
    VTInvalidateWordImageFontSizeControl
    Exit Sub

ApplyFailed:
    applyErrorNumber = Err.Number
    applyErrorDescription = Err.Description
    If internalMutationStarted Then VTEndWordInternalMutation
    Err.Raise applyErrorNumber, "VisualTeX image font size", _
        applyErrorDescription
End Sub

Private Sub VTPrepareWordImageFormulaState( _
    ByVal formulaShape As InlineShape, _
    ByRef formulaId As String, _
    ByRef displayMode As String, _
    ByRef numbered As Boolean, _
    ByRef storedFontSizePt As Double, _
    ByRef referenceWidthPt As Double, _
    ByRef referenceHeightPt As Double, _
    ByRef referenceBaselinePt As Double, _
    ByRef observedWordFontSizePt As Double)

    Dim currentWordFontSizePt As Double
    Dim expectedWidthPt As Double
    Dim expectedHeightPt As Double

    If formulaShape Is Nothing Or VTWordInternalMutationActive() Then Exit Sub
    VTEnsureWordImageScaleState _
        formulaShape, formulaId, displayMode, numbered, storedFontSizePt, _
        referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
        observedWordFontSizePt
    currentWordFontSizePt = VTInlineShapeWordFontSize(formulaShape)
    If VTValidWordFormulaFontSize(currentWordFontSizePt) And _
       Abs(currentWordFontSizePt - observedWordFontSizePt) > 0.05 Then
        ' Selecting an InlineShape can make Word for Mac report the surrounding
        ' paragraph font instead of the VisualTeX size persisted for the image.
        ' The reported value is not tied reliably to the document's Normal style:
        ' Word can return 10 pt for an 11 pt formula, or the reverse. Geometry is
        ' authoritative here. If width and height still represent the persisted
        ' VisualTeX point size, selection changed only Word's transient font view
        ' and must never resize the image.
        expectedWidthPt = referenceWidthPt * _
            storedFontSizePt / VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
        expectedHeightPt = referenceHeightPt * _
            storedFontSizePt / VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT
        If Abs(formulaShape.Width - expectedWidthPt) <= 0.5 And _
           Abs(VTWordImageGlyphHeight(formulaShape) - expectedHeightPt) <= 0.5 Then Exit Sub
        VTApplyWordImageFormulaFontSize formulaShape, currentWordFontSizePt
        VTEnsureWordImageScaleState _
            formulaShape, formulaId, displayMode, numbered, storedFontSizePt, _
            referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
            observedWordFontSizePt
    End If
End Sub

Public Sub VTSynchronizeWordImageFormulaShape( _
    ByVal formulaShape As InlineShape)

    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim storedFontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim observedWordFontSizePt As Double
    Dim scaleStateBefore As String
    Dim scaleStateAfter As String

    If formulaShape Is Nothing Then Exit Sub

    ' An inline formula keeps its Position until it is created, edited or
    ' resized, or the user runs Update Equation Format. Remember the stored
    ' scale state: when this sync resizes the formula (Word font size,
    ' VisualTeX size, dragging the picture), place it by the rule; merely
    ' selecting it changes nothing.
    If VTTryWordImageTypographyReference( _
       formulaShape, formulaId, displayMode, numbered) Then
        If displayMode = "inline" Then
            VTTryGetDocumentVariable formulaShape.Range.Document, _
                VTWordImageScaleVariableName(formulaId), scaleStateBefore
        End If
    End If

    VTPrepareWordImageFormulaState _
        formulaShape, formulaId, displayMode, numbered, storedFontSizePt, _
        referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
        observedWordFontSizePt

    If displayMode = "inline" Then
        VTTryGetDocumentVariable formulaShape.Range.Document, _
            VTWordImageScaleVariableName(formulaId), scaleStateAfter
        If scaleStateAfter <> scaleStateBefore And referenceHeightPt > 0# Then
            VTApplyWordInlineImageBaseline _
                formulaShape, referenceHeightPt, referenceBaselinePt, storedFontSizePt
        End If
    ElseIf displayMode = "block" And Not numbered Then
        VTApplyUnnumberedImageFormulaVerticalAlignment _
            formulaShape.Range
    End If
End Sub

Public Function VTImageTypographyFingerprint(ByVal contextRange As Range) As String
    Dim paragraph As Paragraph
    Dim paragraphMark As Range
    Dim shape As InlineShape
    Dim result As String
    result = contextRange.Document.FullName
    For Each paragraph In contextRange.Paragraphs
        If paragraph.Range.InlineShapes.Count > 0 Then
            result = result & "|" & CStr(paragraph.Range.Start) & ":" & _
                CStr(paragraph.Range.ParagraphFormat.BaseLineAlignment)
            Set paragraphMark = contextRange.Document.Range( _
                Start:=paragraph.Range.End - 1, End:=paragraph.Range.End)
            result = result & ":" & CStr(paragraphMark.Font.Position) & ":" & _
                CStr(paragraphMark.Font.Superscript) & ":" & CStr(paragraphMark.Font.Subscript)
            For Each shape In paragraph.Range.InlineShapes
                result = result & ":" & shape.Title & ":" & _
                    CStr(shape.Height) & ":" & CStr(shape.Width) & ":" & _
                    CStr(shape.Range.Font.Position) & ":" & _
                    Left$(shape.AlternativeText, 64) & ":" & _
                    CStr(shape.Range.Font.Superscript) & ":" & CStr(shape.Range.Font.Subscript)
            Next shape
        End If
    Next paragraph
    VTImageTypographyFingerprint = result
End Function

' Word for Mac can fail to report the font size of a selection that holds a
' picture and field codes; the fingerprint then just lacks it.
Public Function VTSelectionFontSizeText(ByVal selected As Selection) As String
    On Error Resume Next
    VTSelectionFontSizeText = CStr(selected.Font.Size)
    Err.Clear
End Function

Public Function VTDocumentImageTypographyFingerprint(ByVal documentObject As Document) As String
    Dim shape As InlineShape
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim result As String
    For Each shape In documentObject.InlineShapes
        If VTTryWordImageTypographyReference(shape, formulaId, displayMode, numbered) Then
            result = result & "|" & shape.Title & ":" & _
                CStr(shape.Range.ParagraphFormat.BaseLineAlignment) & ":" & _
                CStr(shape.Range.Font.Size) & ":" & CStr(shape.Height) & ":" & _
                CStr(shape.Range.Font.Position) & ":" & CStr(shape.Width) & ":" & _
                CStr(shape.Range.Font.Superscript) & ":" & CStr(shape.Range.Font.Subscript)
        End If
    Next shape
    If Len(result) > 0 Then VTDocumentImageTypographyFingerprint = _
        documentObject.FullName & result
End Function

Public Sub VTEnsureImageSizeWatchScheduled()
    Dim formulaShape As InlineShape
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim hasImages As Boolean

    On Error GoTo ScheduleFailed
    If VT_WORD_IMAGE_SIZE_WATCH_SCHEDULED Or _
       VT_WORD_IMAGE_SIZE_WATCH_RUNNING Or _
       Documents.Count = 0 Or VTWordInternalMutationActive() Then Exit Sub
#If VT_WORD_DIAGNOSTICS Then
    If VT_WORD_TYPOGRAPHY_REGRESSION_RUNNING Then Exit Sub
#End If
    ' Scheduling only needs presence, not another full typography scan.
    For Each formulaShape In ActiveDocument.InlineShapes
        If VTTryWordImageTypographyReference(formulaShape, formulaId, displayMode, numbered) Then
            hasImages = True
            Exit For
        End If
    Next formulaShape
    If Not hasImages Then Exit Sub

    VT_WORD_IMAGE_SIZE_WATCH_SCHEDULED = True
    VT_WORD_IMAGE_WATCH_LAST_DELAY = VT_WORD_IDLE_WATCH_SECONDS
    If Not VTVisualTeXInlineShapeAtSelection(Selection, False) Is Nothing Then _
        VT_WORD_IMAGE_WATCH_LAST_DELAY = VT_WORD_SELECTED_IMAGE_WATCH_SECONDS
    VTScheduleWordIdleTasks VT_WORD_IMAGE_WATCH_LAST_DELAY
    Exit Sub

ScheduleFailed:
    VT_WORD_IMAGE_SIZE_WATCH_SCHEDULED = False
End Sub

Public Function VTWordImageFormulaFontSize( _
    ByVal formulaShape As InlineShape) As Double

    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim observedWordFontSizePt As Double

    If formulaShape Is Nothing Then Exit Function
    If Not VTTryWordImageTypographyReference(formulaShape, formulaId, displayMode, numbered) Then Exit Function
    If VTTryReadWordImageScaleState( _
       formulaShape.Range.Document, formulaId, fontSizePt, referenceWidthPt, _
       referenceHeightPt, referenceBaselinePt, observedWordFontSizePt) Then
        VTWordImageFormulaFontSize = fontSizePt
    Else
        ' A Ribbon getter may report a legacy size, but must not persist it or
        ' adopt geometry. The selection/edit synchronization owns those writes.
        VTWordImageFormulaFontSize = VTPreferredWordFormulaFontSize(formulaShape.Range)
    End If
End Function

Public Function VTWordSelectedImageFormulaState( _
    ByRef formulaCount As Long, _
    ByRef mixedSizes As Boolean, _
    ByRef fontSizePt As Double) As Boolean

    Dim candidate As InlineShape
    Dim adjacentFormula As InlineShape
    Dim candidateSizePt As Double
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean

    formulaCount = 0
    mixedSizes = False
    fontSizePt = 0#
    On Error GoTo StateFailed
    If Documents.Count = 0 Then Exit Function

    If Selection.InlineShapes.Count > 0 Then
        For Each candidate In Selection.InlineShapes
            If VTTryWordImageTypographyReference(candidate, formulaId, displayMode, numbered) Then
                candidateSizePt = VTWordImageFormulaFontSize(candidate)
                formulaCount = formulaCount + 1
                If formulaCount = 1 Then
                    fontSizePt = candidateSizePt
                ElseIf Abs(candidateSizePt - fontSizePt) > 0.05 Then
                    mixedSizes = True
                End If
            End If
        Next candidate
    Else
        Set adjacentFormula = VTVisualTeXInlineShapeAtSelection(Selection, False)
        If Not adjacentFormula Is Nothing Then
            formulaCount = 1
            fontSizePt = VTWordImageFormulaFontSize(adjacentFormula)
        End If
    End If

    VTWordSelectedImageFormulaState = (formulaCount > 0)
    Exit Function

StateFailed:
    formulaCount = 0
    mixedSizes = False
    fontSizePt = 0#
    Err.Clear
End Function

Public Sub VTApplyWordImageFontSizePresetToSelection( _
    ByVal requestedFontSizePt As Double)

    Dim candidate As InlineShape
    Dim adjacentFormula As InlineShape
    Dim appliedCount As Long

    If Not VTValidWordFormulaFontSize(requestedFontSizePt) Then
        Err.Raise vbObjectError + 7498, "VisualTeX", _
            "The selected image formula font-size preset is invalid."
    End If

    If Selection.InlineShapes.Count > 0 Then
        ' One request for the letter fonts their numbers take.
        Dim formulaShapes As New Collection
        For Each candidate In Selection.InlineShapes
            If VTIsVisualTeXInlineShape(candidate) Then formulaShapes.Add candidate
        Next candidate
        VTEnsureWordImageLetterFonts ActiveDocument, formulaShapes
        For Each candidate In Selection.InlineShapes
            If VTIsVisualTeXInlineShape(candidate) Then
                VTApplyWordImageFormulaFontSize candidate, requestedFontSizePt
                appliedCount = appliedCount + 1
            End If
        Next candidate
    Else
        Set adjacentFormula = VTVisualTeXInlineShapeAtSelection(Selection)
        If Not adjacentFormula Is Nothing Then
            VTApplyWordImageFormulaFontSize _
                adjacentFormula, requestedFontSizePt
            adjacentFormula.Select
            appliedCount = 1
        End If
    End If

    If appliedCount = 0 Then
        Err.Raise vbObjectError + 7497, "VisualTeX", _
            "Select one or more VisualTeX image formulas before changing their font size."
    End If
    VTInvalidateWordImageFontSizeControl
End Sub

Public Function VTTryParseWordFormulaFontSize( _
    ByVal enteredText As String, _
    ByRef parsedValue As Double) As Boolean

    Dim normalizedText As String
    Dim index As Long
    Dim currentCharacter As String
    Dim decimalCount As Long
    Dim digitCount As Long

    parsedValue = 0#
    normalizedText = Replace$(Trim$(enteredText), ",", ".")
    If Len(normalizedText) = 0 Then Exit Function
    For index = 1 To Len(normalizedText)
        currentCharacter = Mid$(normalizedText, index, 1)
        If currentCharacter >= "0" And currentCharacter <= "9" Then
            digitCount = digitCount + 1
        ElseIf currentCharacter = "." Then
            decimalCount = decimalCount + 1
            If decimalCount > 1 Then Exit Function
        Else
            Exit Function
        End If
    Next index
    If digitCount = 0 Then Exit Function
    parsedValue = Val(normalizedText)
    VTTryParseWordFormulaFontSize = VTValidWordFormulaFontSize(parsedValue)
End Function

Public Function VTInlineShapeTraceSummary( _
    ByVal selectedShape As InlineShape) As String

    Dim titleText As String
    Dim alternativeText As String

    On Error GoTo InvalidShape
    If selectedShape Is Nothing Then
        VTInlineShapeTraceSummary = "shape=nothing"
        Exit Function
    End If
    titleText = selectedShape.Title
    alternativeText = selectedShape.AlternativeText
    VTInlineShapeTraceSummary = _
        "shapeStart=" & CStr(selectedShape.Range.Start) & _
        " shapeEnd=" & CStr(selectedShape.Range.End) & _
        " shapeType=" & CStr(selectedShape.Type) & _
        " titleLen=" & CStr(Len(titleText)) & _
        " altLen=" & CStr(Len(alternativeText)) & _
        " title=" & VTDoubleClickTraceValue(titleText) & _
        " alt=" & VTDoubleClickTraceValue(alternativeText)
    Exit Function

InvalidShape:
    VTInlineShapeTraceSummary = _
        "shape=error error=" & CStr(Err.Number) & ":" & _
        VTDoubleClickTraceValue(Err.Description)
End Function

' An image formula converted from a display equation keeps the equation's
' math property on its paragraph mark: Word reports an empty OMath there, and
' the tabs and fields of a number inserted anywhere in the paragraph become
' math. OMath.Remove and Font.Reset leave such a mark as it is, so a new mark
' after the formula ends the paragraph and the old mark's empty paragraph is
' deleted (Word keeps the document's final mark).
Public Sub VTClearImageParagraphMathMark(ByVal formulaShape As InlineShape)
    Dim doc As Document
    Dim container As Range
    Dim trailing As Range

    If VTRangeHasMeaningfulMath(formulaShape.Range.Paragraphs(1).Range) Then Exit Sub
    Set doc = formulaShape.Range.Document
    Set container = VTVisualTeXImageContainerRange(formulaShape)
    doc.Range(container.End, container.End).InsertParagraphAfter
    ' The new mark takes the picture's run formatting.
    On Error Resume Next
    doc.Range(container.End, container.End + 1).NoProofing = False
    On Error GoTo 0
    Set trailing = doc.Range(container.End + 1, container.End + 2).Paragraphs(1).Range.Duplicate
    If trailing.End < doc.Content.End And _
       Not VTWordRangeHasMeaningfulText(trailing) Then trailing.Delete
End Sub

Public Sub VTV2ApplyEquationTypography( _
    ByVal exact As Range, ByVal fontSizePt As Double)

    Dim size As Single

    If exact Is Nothing Or fontSizePt <= 0# Then Exit Sub
    ' FormulaFontSize.NormalizeWordOmmlSize: Word keeps half points.
    size = CSng(Int(fontSizePt * 2# + 0.5) / 2#)
    exact.Font.Position = 0
    exact.Font.Size = size
    On Error Resume Next
    exact.Font.SizeBi = size
    Err.Clear
End Sub

' Inline image formulas whose Position differs from VTWordInlineImagePosition
' (formulas made before r117, or moved by hand). With apply, place each one by
' the rule. Formulas without readable scale state are left alone.
Public Function VTV2InlineBaselineRepairs( _
    ByVal doc As Document, ByVal apply As Boolean) As Long

    Dim shape As InlineShape
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim observedWordFontSizePt As Double
    Dim expected As Long
    Dim savedCanvas As String
    Dim readable As Boolean
    Dim needsRepair As Boolean

    For Each shape In doc.InlineShapes
        needsRepair = False
        If VTTryWordImageTypographyReference(shape, formulaId, displayMode, numbered) Then
            If displayMode = "inline" Then
                On Error Resume Next
                readable = VTTryReadWordImageScaleState(doc, formulaId, fontSizePt, _
                    referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
                    observedWordFontSizePt)
                If Err.Number <> 0 Then readable = False
                Err.Clear
                If readable Then
                    expected = VTWordInlineImagePosition( _
                        shape, referenceHeightPt, referenceBaselinePt, fontSizePt)
                    If Err.Number <> 0 Then readable = False
                    Err.Clear
                End If
                On Error GoTo 0
                If readable Then
                    needsRepair = shape.Range.Font.Position <> expected Or _
                        shape.Range.Font.Superscript Or shape.Range.Font.Subscript
                    If Not needsRepair Then
                        If VTTryGetDocumentVariable(doc, _
                           VTWordImageCanvasVariableName(formulaId), savedCanvas) Then _
                            needsRepair = VTCachedWordImageCropTop(shape) < 0#
                    End If
                End If
            End If
        End If
        If needsRepair Then
            VTV2InlineBaselineRepairs = VTV2InlineBaselineRepairs + 1
            If apply Then VTApplyWordInlineImageBaseline _
                shape, referenceHeightPt, referenceBaselinePt, fontSizePt
        End If
    Next shape
End Function
