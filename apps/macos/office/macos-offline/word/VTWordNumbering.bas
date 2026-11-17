Attribute VB_Name = "VTWordNumbering"
Option Explicit
Option Private Module

' Equation numbers: SEQ fields and their bookmarks, captions, number helpers and their repair, and equation
' format v2 numbers.

Public Const VT_WORD_NATIVE_NUMBER_CACHE_PLACEHOLDER As String = _
    "VTNUMBERCACHE0"

Public Function VTNumberedDisplayTableNearRange( _
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

Public Function VTNumberBookmarkNameForTable( _
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

Public Function VTFormulaIdFromBookmarkSuffix( _
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

Public Function VTFormulaIdFromSequenceBookmarkName( _
    ByVal bookmarkName As String) As String

    If Left$(bookmarkName, _
       Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)) <> _
       VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX Then Exit Function
    VTFormulaIdFromSequenceBookmarkName = VTFormulaIdFromBookmarkSuffix( _
        Mid$(bookmarkName, _
            Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX) + 1))
End Function

Public Function VTSequenceBookmarkNameForField( _
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

Public Function VTFormulaIdForValidNumberedTable( _
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

Public Function VTNumberedFormulaRangeForId( _
    ByVal documentObject As Document, _
    ByVal formulaId As String) As Range

    Dim numberRange As Range
    Dim formulaContainer As Range
    Dim repairedNumberRange As Range
    Dim nativeRange As Range
    Dim layoutTable As Table
    Dim nativeMath As OMath
    Dim imageShape As InlineShape
    Dim visibleNumberField As Field
    Dim numberBookmarkName As String
    Dim nativeBookmarkName As String
    Dim displayMode As String
    Dim encodedMetadata As String
    Dim formulaReference As String
    Dim numbered As Boolean

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
                    VTNativeMathRoot(formulaContainer.OMaths(1)).Range.Duplicate
                Exit Function
            End If
        End If
    End If

    ' During a newly inserted image formula transaction, Word for Mac can
    ' invalidate or temporarily detach the freshly-created VT_R_ Range while a
    ' neighboring SEQ field is updated. The image itself already has a stronger
    ' durable identity at this point: formulaId + stored metadata + canonical
    ' Title. Resolve that exact object before falling back to native OMath. Never
    ' use a nearest-image scan, because that could bind a new number to a
    ' neighboring formula.
    If VTTryReadWordFormulaFormat( _
       documentObject, formulaId, displayMode, numbered) Then
        If displayMode = "block" And numbered Then
            If VTTryReadWordMetadataPayload( _
               documentObject, formulaId, encodedMetadata) Then
                If VTIsEncodedMetadata(encodedMetadata) Then
                    formulaReference = _
                        VTFormulaReference(formulaId, displayMode, numbered)
                    Set imageShape = VTFindCommittedInlineShapeInDocument( _
                        documentObject, encodedMetadata, formulaReference)
                    If Not imageShape Is Nothing Then
                        If Not imageShape.Range.Information(wdWithInTable) Then
                            Set VTNumberedFormulaRangeForId = _
                                imageShape.Range.Duplicate
                            Exit Function
                        End If
                    End If
                End If
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
    If nativeRange.Information(wdWithInTable) Then
        If nativeRange.Tables.Count <> 1 Then Exit Function
        Set layoutTable = nativeRange.Tables(1)
        If layoutTable.Rows.Count <> 1 Or layoutTable.Columns.Count <> 3 Then
            Exit Function
        End If
        Set formulaContainer = layoutTable.Cell(1, 2).Range.Duplicate
        If formulaContainer.InlineShapes.Count <> 0 Or _
           formulaContainer.OMaths.Count <> 1 Then
            Exit Function
        End If
        If nativeRange.Start < formulaContainer.Start Or _
           nativeRange.End > formulaContainer.End Then
            Exit Function
        End If
        Set VTNumberedFormulaRangeForId = nativeRange.Duplicate
        Exit Function
    End If

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

Public Function VTValidNumberedFormulaIds( _
    ByVal documentObject As Document) As Variant
    Dim shape As InlineShape
    Dim bookmark As Bookmark
    Dim ids As New Collection
    Dim result() As String
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim index As Long
    If documentObject Is Nothing Then Exit Function
    For Each shape In documentObject.InlineShapes
        If VTTryParseFormulaReference(shape.Title, formulaId, displayMode, numbered) Then
            If numbered And displayMode = "block" Then
                If Not VTImageEquationReferenceField(shape.Range, formulaId) Is Nothing Then
                    VTCollectionAddUniqueText ids, formulaId
                End If
            End If
        End If
    Next shape
    ' Retain legacy native identities only until their next explicit migration.
    For Each bookmark In documentObject.Bookmarks
        If VTTryFormulaIdFromNativeBookmark(bookmark.Name, formulaId) Then
            If documentObject.Bookmarks.Exists(VTEquationNumberBookmarkName(formulaId)) Then
                VTCollectionAddUniqueText ids, formulaId
            End If
        End If
    Next bookmark
    If ids.Count = 0 Then Exit Function
    ReDim result(1 To ids.Count)
    For index = 1 To ids.Count
        result(index) = CStr(ids(index))
    Next index
    VTValidNumberedFormulaIds = result
End Function

Public Function VTParagraphHasNumberHelperGeometry( _
    ByVal paragraphRange As Range) As Boolean

    Dim helperFrame As Frame
    Dim frameHorizontalPosition As Single
    Dim frameVerticalPosition As Single
    Dim pageWidth As Single
    Dim pageHeight As Single

    If paragraphRange Is Nothing Or _
       paragraphRange.Information(wdWithInTable) Or _
       paragraphRange.InlineShapes.Count <> 0 Or _
       paragraphRange.OMaths.Count <> 0 Then Exit Function

    ' Current helpers match the mature Windows production Frame exactly: a
    ' full-size native SEQ caption clipped inside a 0.1-point square at the
    ' bottom-right page boundary. Word then keeps its private _Ref on the pure
    ' helper caption instead of the preceding tab/image/visible-REF formula row.
    ' Keep both prior Mac geometries as migration signatures so schema 8 can
    ' still recognize and repair documents created before this production Frame.
    If paragraphRange.Frames.Count = 1 Then
        Set helperFrame = paragraphRange.Frames(1)
        With paragraphRange.ParagraphFormat
            If .Alignment <> wdAlignParagraphLeft Or _
               Abs(.LeftIndent) > 0.2 Or _
               Abs(.RightIndent) > 0.2 Or _
               Abs(.FirstLineIndent) > 0.2 Or _
               Abs(.SpaceBefore) > 0.2 Or _
               Abs(.SpaceAfter) > 0.2 Then Exit Function
        End With
        With helperFrame
            If .HeightRule <> wdFrameExact Or _
               Abs(.Height - 0.1) > 0.05 Or _
               .WidthRule <> wdFrameExact Or _
               Abs(.Width - 0.1) > 0.05 Or _
               .RelativeHorizontalPosition <> _
                   wdRelativeHorizontalPositionPage Or _
               .RelativeVerticalPosition <> _
                   wdRelativeVerticalPositionPage Or _
               .TextWrap Then Exit Function
        End With

        If paragraphRange.Sections.Count < 1 Then Exit Function
        pageWidth = paragraphRange.Sections(1).PageSetup.PageWidth
        pageHeight = paragraphRange.Sections(1).PageSetup.PageHeight
        frameHorizontalPosition = 0!
        frameVerticalPosition = 0!
        If pageWidth > 0.1 Then _
            frameHorizontalPosition = pageWidth - 0.1
        If pageHeight > 0.1 Then _
            frameVerticalPosition = pageHeight - 0.1

        If paragraphRange.ParagraphFormat.LineSpacingRule = _
               wdLineSpaceSingle And _
           Abs(helperFrame.HorizontalPosition - _
               frameHorizontalPosition) <= 0.2 And _
           Abs(helperFrame.VerticalPosition - _
               frameVerticalPosition) <= 0.2 Then
            VTParagraphHasNumberHelperGeometry = True
            Exit Function
        End If

        ' r88 migration input: the first real-Frame implementation used a
        ' top-left Frame and an exact one-point line. It is recognizable here
        ' only so normal reconciliation can move it to the production geometry.
        If paragraphRange.ParagraphFormat.LineSpacingRule = _
               wdLineSpaceExactly And _
           Abs(paragraphRange.ParagraphFormat.LineSpacing - 1!) <= 0.2 And _
           helperFrame.HorizontalPosition = wdFrameLeft And _
           helperFrame.VerticalPosition = wdFrameTop Then
            VTParagraphHasNumberHelperGeometry = True
        End If
        Exit Function
    ElseIf paragraphRange.Frames.Count > 1 Then
        Exit Function
    End If

    ' Pre-Frame schema-7 migration input.
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

    VTParagraphHasNumberHelperGeometry = True
End Function

Public Function VTHelperParagraphOwnsNativeEquationSequence( _
    ByVal paragraphRange As Range) As Boolean

    Dim candidateField As Field

    ' VBA's Or evaluates every operand, and Range.Information takes about 5 ms
    ' in Word for Mac (measured): the cheap tests come first and stop here.
    If paragraphRange Is Nothing Then Exit Function
    If paragraphRange.Fields.Count <> 1 Then Exit Function
    If paragraphRange.OMaths.Count <> 0 Then Exit Function
    If paragraphRange.InlineShapes.Count <> 0 Then Exit Function
    If paragraphRange.Information(wdWithInTable) Then Exit Function
    Set candidateField = paragraphRange.Fields(1)
    VTHelperParagraphOwnsNativeEquationSequence = _
        VTIsNativeEquationSequenceField( _
            candidateField, VTNativeEquationLabelName())
End Function

Public Function VTIsDetachedVisualTeXNativeSequenceHelper( _
    ByVal paragraphRange As Range) As Boolean

    Dim candidateField As Field
    Dim sequenceField As Field
    Dim sequenceCount As Long

    If paragraphRange Is Nothing Or _
       Not VTParagraphHasNumberHelperGeometry( _
           paragraphRange) Then Exit Function

    ' A healthy helper contains one managed SEQ. A reference-selection drift can
    ' append a body REF to that same compact paragraph, so Fields.Count can be 2
    ' (or more) after corruption. Count the SEQ identity independently and allow
    ' extra non-SEQ fields; the exact helper geometry above remains the ownership
    ' signature and prevents ordinary Word captions from entering this cleanup.
    For Each candidateField In paragraphRange.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            sequenceCount = sequenceCount + 1
            Set sequenceField = candidateField
        End If
    Next candidateField
    If sequenceCount <> 1 Or sequenceField Is Nothing Then Exit Function

    ' A live VisualTeX helper always owns a VT_N_ Bookmark. Once that durable
    ' identity is gone, this paragraph is detached even when an accidental REF
    ' has polluted it. Explicit orphan repair may then delete the whole helper.
    If Len(VTSequenceBookmarkNameForField( _
       paragraphRange.Document, sequenceField)) > 0 Then Exit Function

    VTIsDetachedVisualTeXNativeSequenceHelper = True
End Function

Public Function VTDetachedVisualTeXNativeSequenceHelperNearRange( _
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

Private Function VTFormulaIdForDetachedSequenceHelper( _
    ByVal documentObject As Document, _
    ByVal helperParagraph As Range) As String

    Dim candidateBookmark As Bookmark
    Dim precedingParagraph As Range
    Dim formulaId As String
    Dim candidateId As String
    Dim matchCount As Long

    If documentObject Is Nothing Or helperParagraph Is Nothing Then Exit Function

    ' VT_C_ normally survives when Word drops only the field-result VT_N_. It is
    ' the strongest recovery identity because it belongs to the helper paragraph
    ' itself. A visible VT_R_ in the immediately preceding formula paragraph is
    ' the fallback for partially deleted or cut formulas.
    For Each candidateBookmark In documentObject.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_CAPTION_BOOKMARK_PREFIX)) = _
           VT_WORD_CAPTION_BOOKMARK_PREFIX Then
            If candidateBookmark.Range.Start >= helperParagraph.Start And _
               candidateBookmark.Range.Start <= helperParagraph.End Then
                candidateId = VTFormulaIdFromBookmarkSuffix( _
                    Mid$(candidateBookmark.Name, _
                        Len(VT_WORD_CAPTION_BOOKMARK_PREFIX) + 1))
                If Len(candidateId) > 0 Then
                    matchCount = matchCount + 1
                    formulaId = candidateId
                End If
            End If
        End If
    Next candidateBookmark
    If matchCount = 1 Then
        VTFormulaIdForDetachedSequenceHelper = formulaId
        Exit Function
    ElseIf matchCount > 1 Or helperParagraph.Start <= 0 Then
        Exit Function
    End If

    Set precedingParagraph = documentObject.Range( _
        Start:=helperParagraph.Start - 1, _
        End:=helperParagraph.Start - 1).Paragraphs(1).Range.Duplicate
    For Each candidateBookmark In documentObject.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_NUMBER_BOOKMARK_PREFIX Then
            If candidateBookmark.Range.Start >= precedingParagraph.Start And _
               candidateBookmark.Range.End <= precedingParagraph.End Then
                candidateId = VTFormulaIdFromBookmarkSuffix( _
                    Mid$(candidateBookmark.Name, _
                        Len(VT_WORD_NUMBER_BOOKMARK_PREFIX) + 1))
                If Len(candidateId) > 0 Then
                    matchCount = matchCount + 1
                    formulaId = candidateId
                End If
            End If
        End If
    Next candidateBookmark
    If matchCount = 1 Then _
        VTFormulaIdForDetachedSequenceHelper = formulaId
End Function

Public Function VTPruneDetachedVisualTeXNativeSequenceHelpers( _
    ByVal documentObject As Document) As Long

    Dim candidateField As Field
    Dim helperParagraph As Range
    Dim formulaRange As Range
    Dim sequenceField As Field
    Dim helperStarts() As Long
    Dim helperFormulaIds() As String
    Dim helperCount As Long
    Dim formulaId As String
    Dim itemIndex As Long

    If documentObject Is Nothing Then Exit Function

    ' First capture stable paragraph anchors without mutating Fields. Word for
    ' Mac can throw 5941 or skip later fields when a For Each Fields traversal
    ' deletes or rebuilds a field in the same collection.
    For Each candidateField In documentObject.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            ' Managed owners are validated by the following live-formula pass.
            ' Avoid expensive frame/font geometry queries for every healthy SEQ.
            If Len(VTSequenceBookmarkNameForField( _
               documentObject, candidateField)) = 0 Then
            Set helperParagraph = _
                candidateField.Result.Paragraphs(1).Range.Duplicate
            If VTIsDetachedVisualTeXNativeSequenceHelper( _
               helperParagraph) Then
                helperCount = helperCount + 1
                If helperCount = 1 Then
                    ReDim helperStarts(1 To 1)
                    ReDim helperFormulaIds(1 To 1)
                Else
                    ReDim Preserve helperStarts(1 To helperCount)
                    ReDim Preserve helperFormulaIds(1 To helperCount)
                End If
                helperStarts(helperCount) = helperParagraph.Start
                helperFormulaIds(helperCount) = _
                    VTFormulaIdForDetachedSequenceHelper( _
                        documentObject, helperParagraph)
            End If
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
                formulaId = helperFormulaIds(itemIndex)
                Set formulaRange = Nothing
                If Len(formulaId) > 0 Then
                    Set formulaRange = VTNumberedFormulaRangeForId( _
                        documentObject, formulaId)
                End If
                If Not formulaRange Is Nothing Then
                    Set sequenceField = Nothing
                    For Each candidateField In helperParagraph.Fields
                        If VTIsNativeEquationSequenceField( _
                           candidateField, VTNativeEquationLabelName()) Then
                            Set sequenceField = candidateField
                            Exit For
                        End If
                    Next candidateField
                    If sequenceField Is Nothing Then
                        Err.Raise vbObjectError + 7549, "VisualTeX", _
                            "A live Equation helper lost its SEQ field during repair."
                    End If
                    If documentObject.Bookmarks.Exists( _
                       VTEquationSequenceNumberBookmarkName(formulaId)) Then
                        documentObject.Bookmarks( _
                            VTEquationSequenceNumberBookmarkName(formulaId)).Delete
                    End If
                    documentObject.Bookmarks.Add _
                        name:=VTEquationSequenceNumberBookmarkName(formulaId), _
                        Range:=sequenceField.Result.Duplicate
                    VTSetCollapsedEquationCaptionBookmark _
                        documentObject, formulaId, helperParagraph
                    VTFormatHiddenEquationParagraph helperParagraph
                Else
                    ' Delete the exact helper first. Removing VT_R_ can shift the
                    ' following paragraph, so deleting by its freshly verified
                    ' Range avoids depending on the old helper anchor.
                    helperParagraph.Delete
                    If Len(formulaId) > 0 Then
                        VTDeleteEquationNumberScaffold _
                            documentObject, formulaId, True
                    End If
                End If
                VTPruneDetachedVisualTeXNativeSequenceHelpers = _
                    VTPruneDetachedVisualTeXNativeSequenceHelpers + 1
            End If
        End If
    Next itemIndex
End Function

Public Sub VTDeleteEquationNumberScaffold( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    Optional ByVal deleteTable As Boolean = True, _
    Optional ByVal preserveFormulaState As Boolean = False)

    Dim numberBookmarkName As String
    Dim sequenceBookmarkName As String
    Dim captionBookmarkName As String
    Dim nativeBookmarkName As String
    Dim expectedNumberText As String
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
    ' Prefer the exact SEQ field result when resolving the hidden helper.
    ' VT_C_ is a collapsed Bookmark at the end of that helper paragraph; on
    ' Word for Mac, Paragraphs(1) from a zero-length paragraph-end Range can
    ' resolve to the following body paragraph. That left the real SEQ helper
    ' behind after formula-to-LaTeX restoration even though VT_N_/VT_C_ were
    ' already deleted. The VT_N_ result is the durable helper identity.
    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        Set sequenceField = VTEquationSequenceFieldForBookmark( _
            documentObject, sequenceBookmarkName)
        If Not sequenceField Is Nothing Then
            Set sequenceHelperParagraph = _
                sequenceField.Result.Paragraphs(1).Range.Duplicate
        End If
    End If
    If sequenceHelperParagraph Is Nothing And _
       documentObject.Bookmarks.Exists(captionBookmarkName) Then
        Set sequenceHelperParagraph = documentObject.Bookmarks( _
            captionBookmarkName).Range.Paragraphs(1).Range.Duplicate
    End If

    ' Remove the structural identity and helper before touching body REF fields.
    ' Rollback runs this routine under On Error Resume Next; if a damaged body
    ' reference raises first, VBA abandons the rest of this procedure and used
    ' to leave VT_N_/VT_C_ plus the SEQ helper behind. That orphan then poisoned
    ' the next numbering reconciliation. Make the formula-specific scaffold the
    ' first cleanup responsibility so a later reference-repair failure cannot
    ' strand another managed SEQ field.
    If documentObject.Bookmarks.Exists(numberBookmarkName) Then
        documentObject.Bookmarks(numberBookmarkName).Delete
    End If
    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        documentObject.Bookmarks(sequenceBookmarkName).Delete
    End If
    If documentObject.Bookmarks.Exists(captionBookmarkName) Then
        documentObject.Bookmarks(captionBookmarkName).Delete
    End If
    If Not preserveFormulaState And _
       documentObject.Bookmarks.Exists(nativeBookmarkName) Then
        documentObject.Bookmarks(nativeBookmarkName).Delete
    End If
    If Not sequenceHelperParagraph Is Nothing Then
        If VTHelperParagraphOwnsNativeEquationSequence( _
           sequenceHelperParagraph) Then
            sequenceHelperParagraph.Delete
        End If
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
    If deleteTable And Not layoutTable Is Nothing Then layoutTable.Delete

    ' Turning numbering off keeps the formula's edit identity and payload.
    If preserveFormulaState Then Exit Sub

    ' Remove the formula's own visible number before repairing any remaining
    ' body REF fields. Otherwise the number REF is mistaken for a cross-reference
    ' and restored LaTeX can be followed by a "(deleted equation)" marker.
    VTReplaceDeletedEquationBodyReferences _
        documentObject, sequenceBookmarkName

    VTDeleteWordLatexPayload documentObject, formulaId
    VTDeleteWordOmmlPayload documentObject, formulaId
    VTDeleteWordMetadataPayload documentObject, formulaId
    VTDeleteDocumentVariable _
        documentObject, VTWordFormatVariableName(formulaId)
    VTDeleteDocumentVariable _
        documentObject, VTWordImageScaleVariableName(formulaId)
    VTDeleteDocumentVariable documentObject, VTWordImageCanvasVariableName(formulaId)
    VTDeleteDocumentVariable _
        documentObject, VTWordImageInkCenterVariableName(formulaId)
    VTDeleteDocumentVariable _
        documentObject, VTWordImageLetterFontVariableName(formulaId)
    VTDeleteDocumentVariable _
        documentObject, VTWordNativeSignatureVariableName(formulaId)
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

Public Function VTPruneOrphanedEquationNumberScaffolds( _
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

    ' Reuse identities while the document is unchanged. After a deletion all
    ' Word positions must be resolved again before the next field traversal.
    If orphanIds.Count > 0 Then
        validFormulaIds = VTValidNumberedFormulaIds(documentObject)
        Set validIds = New Collection
        If IsArray(validFormulaIds) Then
            For itemIndex = LBound(validFormulaIds) To UBound(validFormulaIds)
                VTCollectionAddUniqueText validIds, CStr(validFormulaIds(itemIndex))
            Next itemIndex
        End If
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

    If documentObject.Tables.Count > 0 Then
        VTRepairLiveNumberedTableScaffolds documentObject
        validFormulaIds = VTValidNumberedFormulaIds(documentObject)
    End If
    VTPruneOrphanedEquationNumberScaffolds = _
        VTVariantArrayCount(validFormulaIds)
End Function

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

Public Sub VTEnsureOrphanWatchScheduled()
    On Error GoTo ScheduleFailed
    If VT_WORD_ORPHAN_WATCH_SCHEDULED Or _
       VT_WORD_ORPHAN_WATCH_RUNNING Then Exit Sub
    If Documents.Count = 0 Then Exit Sub
    If Not VTAnyOpenDocumentHasEquationNumbers() Then Exit Sub

    VT_WORD_ORPHAN_WATCH_SCHEDULED = True
    VTScheduleWordIdleTasks VT_WORD_IDLE_WATCH_SECONDS
    Exit Sub

ScheduleFailed:
    VT_WORD_ORPHAN_WATCH_SCHEDULED = False
End Sub

' WPS turns every "*" in equation text into U+2217 ASTERISK OPERATOR when it
' opens a document, the field instructions inside numbered equations included,
' and saves them that way. Word then reads "SEQ Equation \s 1 \<U+2217> ARABIC"
' as a SEQ with a bookmark argument and the number shows "Error! Bookmark not
' defined.". Word makes the same substitution in a field code typed into an
' equation, so the number is rebuilt instead: such an equation still counts as
' numbered (VTV2NumberedEquationEntry) and its number is not in the document's
' style, so VTV2RestyleNumbers writes it again from the shared builder.
Private Function VTIsWpsDamagedNumberInstruction(ByVal instruction As String) As Boolean
    VTIsWpsDamagedNumberInstruction = _
        InStr(1, instruction, "\" & ChrW(&H2217), vbBinaryCompare) > 0
End Function

Public Function VTHasWpsDamagedEquationNumber(ByVal documentObject As Document) As Boolean
    Dim candidate As Field
    Dim fieldType As Long

    If documentObject.OMaths.Count = 0 Then Exit Function
    For Each candidate In documentObject.Fields
        fieldType = candidate.Type
        If fieldType = wdFieldSequence Or fieldType = wdFieldStyleRef Then
            If VTIsWpsDamagedNumberInstruction(candidate.Code.Text) Then
                VTHasWpsDamagedEquationNumber = True
                Exit Function
            End If
        End If
    Next candidate
End Function

' Documents opened before the add-in's event sink existed (the one Word was
' launched with) get the same repair when the Ribbon loads.
Public Sub VTRepairOpenedWpsDocuments()
    Dim documentObject As Document

    On Error Resume Next
    For Each documentObject In Documents
        VTRepairOpenedWpsDocument documentObject
        Err.Clear
    Next documentObject
    On Error GoTo 0
End Sub

Public Function VTUpdateEquationNumbersCore( _
    ByVal documentObject As Document, _
    Optional ByVal performOrphanRepair As Boolean = True) As Long

    Dim movedHelpers As Long
    Dim referenceBindings As Collection
    Dim previousScreenUpdating As Boolean
    Dim mutationStarted As Boolean
    Dim operationErrorNumber As Long
    Dim operationErrorDescription As String

    If documentObject Is Nothing Or VT_WORD_NUMBER_UPDATE_RUNNING Then Exit Function
    On Error GoTo Failed
    VT_WORD_NUMBER_UPDATE_RUNNING = True
    previousScreenUpdating = Application.ScreenUpdating
    VTBeginWordInternalMutation
    mutationStarted = True
    Application.ScreenUpdating = False
    VTUpgradeWordNumberFormats documentObject
    VTUpdateCurrentWordNumberFields documentObject
    VTUpdateEquationNumbersCore = VTCurrentNumberTargets(documentObject).Count
    VTWordPerformanceMark "numbering-complete"

Finished:
    On Error Resume Next
    If mutationStarted Then VTEndWordInternalMutation
    Application.ScreenUpdating = previousScreenUpdating
    VT_WORD_NUMBER_UPDATE_RUNNING = False
    On Error GoTo 0
    If operationErrorNumber <> 0 Then
        Err.Raise operationErrorNumber, "VisualTeX Equation numbering", _
            operationErrorDescription
    End If
    Exit Function

Failed:
    operationErrorNumber = Err.Number
    operationErrorDescription = Err.Description
    Resume Finished
End Function

Public Function VTInsertDedicatedEquationHelperParagraph( _
    ByVal layoutTable As Table) As Range

    Dim documentObject As Document
    Dim tableRange As Range
    Dim probeRange As Range
    Dim helperParagraph As Range
    Dim candidateParagraph As Range
    Dim helperStart As Long
    Dim probeStart As Long
    Dim probeLimit As Long

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

    ' A collapsed Range exactly at the old Table.Range.End can still resolve
    ' backward into the final cell marker (Chr(13) + Chr(7)) on Word 16.89.1.
    ' Probe actual characters immediately after that boundary and select the
    ' first paragraph that Word reports as outside the table. Do not accept a
    ' later body paragraph: the dedicated helper must also be completely empty.
    probeStart = helperStart
    probeLimit = helperStart + 8
    If probeLimit > documentObject.Content.End - 1 Then
        probeLimit = documentObject.Content.End - 1
    End If
    Do While probeStart <= probeLimit
        Set probeRange = documentObject.Range( _
            Start:=probeStart, End:=probeStart + 1)
        Set candidateParagraph = probeRange.Paragraphs(1).Range.Duplicate
        If Not candidateParagraph.Information(wdWithInTable) And _
           candidateParagraph.Start >= helperStart Then
            Set helperParagraph = candidateParagraph
            Exit Do
        End If
        probeStart = probeStart + 1
    Loop

    If helperParagraph Is Nothing Then
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "Word did not expose an Equation helper paragraph outside the table."
    End If
    If helperParagraph.Fields.Count <> 0 Or _
       helperParagraph.InlineShapes.Count <> 0 Or _
       helperParagraph.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(helperParagraph) Then
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "Word did not create an independent empty Equation helper paragraph" & _
            " [expectedStart=" & CStr(helperStart) & _
            "; actualStart=" & CStr(helperParagraph.Start) & _
            "; actualEnd=" & CStr(helperParagraph.End) & _
            "; fields=" & CStr(helperParagraph.Fields.Count) & _
            "; maths=" & CStr(helperParagraph.OMaths.Count) & _
            "; text=" & Replace$(helperParagraph.Text, vbCr, "<CR>") & "]."
    End If
    Set VTInsertDedicatedEquationHelperParagraph = helperParagraph
End Function

Public Function VTInsertEquationNumber( _
    ByRef formulaShape As InlineShape, _
    ByVal formulaId As String, _
    ByVal captionText As String, _
    Optional ByVal deferReconcile As Boolean = False, _
    Optional ByRef performancePhases As Variant) As Range
    If formulaShape.Range.Information(wdWithInTable) Then
        Set VTInsertEquationNumber = VTMigrateLegacyNumberedImageTableToParagraph( _
            formulaShape, formulaId, captionText)
        Exit Function
    End If
    Set VTInsertEquationNumber = VTWriteSingleParagraphImageNumber(formulaShape, formulaId)
    If Not deferReconcile Then
        ' Image and native equations deliberately own different SEQ chains. A
        ' newly inserted image can only change image numbers at or after its
        ' paragraph, so do not rescan every earlier image or touch OMML fields.
        ' Besides preserving that separation, this keeps repeated appends
        ' linear instead of making document construction quadratic.
        VTUpdateImageNumberFieldsFrom _
            formulaShape.Range.Document, formulaShape.Range.Start
    End If
    If Not IsMissing(performancePhases) Then performancePhases = Array(0#, 0#, 0#, 0#, 0#, 0#)
End Function

Public Function VTNativeEquationLabelName() As String
    Dim equationLabel As CaptionLabel

    Set equationLabel = Application.CaptionLabels(wdCaptionEquation)
    VTNativeEquationLabelName = Trim$(equationLabel.Name)
    If Len(VTNativeEquationLabelName) = 0 Then
        Err.Raise vbObjectError + 7429, "VisualTeX", "Word did not expose its built-in Equation caption label."
    End If
End Function

Public Function VTEquationSequenceFieldText(ByVal equationLabelName As String) As String
    ' Return only the localized SEQ identifier. The production formatting switch
    ' is appended by the external helper field builders after legacy direct-SEQ
    ' input has been moved out of OMath; Word for Mac can then retain an ordinary
    ' asterisk without converting it into a mathematical operator.
    If InStr(1, equationLabelName, " ", vbBinaryCompare) > 0 Then
        VTEquationSequenceFieldText = _
            """" & Replace$(equationLabelName, """", """""") & """"
    Else
        VTEquationSequenceFieldText = equationLabelName
    End If
End Function

Public Function VTInsertRegisteredEquationCaption( _
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
    ' Match the reconciled Windows production field exactly. The helper is now
    ' always external to OMath, so Word for Mac preserves the ordinary asterisk
    ' in \* ARABIC. PreserveFormatting remains False only to avoid an additional
    ' MERGEFORMAT switch; the native ARABIC switch itself is canonical.
    Set sequenceField = documentObject.Fields.Add( _
        Range:=insertionRange, _
        Type:=wdFieldEmpty, _
        Text:="SEQ " & VTEquationSequenceFieldText(equationLabelName) & _
            " \* ARABIC", _
        PreserveFormatting:=False)
    ' Word normally evaluates a newly inserted SEQ immediately. Avoid forcing a
    ' second sequence walk on every append; that cost grows with the number of
    ' preceding Equation fields. Keep one compatibility fallback for builds that
    ' leave the fresh result unresolved until Update is requested explicitly.
    If VTFirstPositiveIntegerInText( _
       VTEquationSequenceResultText(sequenceField)) < 1 Then
        sequenceField.Update
    End If

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

Public Function VTIsNativeEquationSequenceField( _
    ByVal candidate As Field, _
    ByVal equationLabelName As String) As Boolean

    Dim fieldCode As String
    Dim sequenceIdentifier As String
    Dim documentObject As Document
    Dim helperRange As Range
    Dim captionRange As Range
    Dim sequenceRange As Range
    Dim sequenceBookmarkName As String
    Dim captionBookmarkName As String
    Dim formulaId As String
    If candidate Is Nothing Then Exit Function
    If candidate.Type <> wdFieldSequence Then Exit Function
    fieldCode = candidate.Code.Text
    sequenceIdentifier = VTSequenceIdentifierFromFieldCode(fieldCode)
    VTIsNativeEquationSequenceField = _
        StrComp(sequenceIdentifier, equationLabelName, vbTextCompare) = 0
    If VTIsNativeEquationSequenceField Then Exit Function

    ' Equation and 公式 are literal SEQ identifiers, so Word does not advance one
    ' merely because the other is the current localized Caption label. Accept a
    ' different identifier only under both exact schema-7 ownership bookmarks;
    ' normal reconciliation then migrates that managed field to the active label.
    Set documentObject = candidate.Result.Document
    sequenceBookmarkName = VTSequenceBookmarkNameForField(documentObject, candidate)
    formulaId = VTFormulaIdFromSequenceBookmarkName(sequenceBookmarkName)
    If Not VTIsCanonicalUuid(formulaId) Then Exit Function
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(captionBookmarkName) Then Exit Function
    Set helperRange = candidate.Result.Paragraphs(1).Range.Duplicate
    Set captionRange = documentObject.Bookmarks(captionBookmarkName).Range.Duplicate
    Set sequenceRange = documentObject.Bookmarks(sequenceBookmarkName).Range.Duplicate
    If helperRange.Information(wdWithInTable) Or _
       helperRange.Fields.Count <> 1 Or helperRange.OMaths.Count <> 0 Or _
       helperRange.InlineShapes.Count <> 0 Then Exit Function
    ' Word for Mac serializes a Range ending at Field.Result.End with the
    ' Bookmark end after the field-end character. After reopen the API therefore
    ' reports VT_N_.End as either Result.End or Result.End + 1. Both are the same
    ' exact field-owned boundary; reject anything beyond VTEquationFieldEnd.
    VTIsNativeEquationSequenceField = _
        captionRange.Start = helperRange.Start And captionRange.End = helperRange.End And _
        sequenceRange.Start >= helperRange.Start And _
        sequenceRange.Start <= candidate.Result.Start And _
        sequenceRange.End >= candidate.Result.End And _
        sequenceRange.End <= VTEquationFieldEnd(candidate)
End Function

Public Function VTSequenceIdentifierFromFieldCode( _
    ByVal fieldCode As String) As String

    Dim normalizedCode As String
    Dim remainder As String
    Dim identifier As String
    Dim characterValue As String
    Dim characterIndex As Long

    normalizedCode = VTNormalizeEquationFieldCode(fieldCode)
    If Len(normalizedCode) < 4 Or _
       StrComp(Left$(normalizedCode, 3), "SEQ", vbTextCompare) <> 0 Or _
       Mid$(normalizedCode, 4, 1) <> " " Then Exit Function
    remainder = LTrim$(Mid$(normalizedCode, 5))
    If Len(remainder) = 0 Then Exit Function

    If Left$(remainder, 1) = """" Then
        characterIndex = 2
        Do While characterIndex <= Len(remainder)
            characterValue = Mid$(remainder, characterIndex, 1)
            If characterValue = """" Then
                If characterIndex < Len(remainder) And _
                   Mid$(remainder, characterIndex + 1, 1) = """" Then
                    identifier = identifier & """"
                    characterIndex = characterIndex + 2
                Else
                    VTSequenceIdentifierFromFieldCode = identifier
                    Exit Function
                End If
            Else
                identifier = identifier & characterValue
                characterIndex = characterIndex + 1
            End If
        Loop
        Exit Function
    End If

    For characterIndex = 1 To Len(remainder)
        characterValue = Mid$(remainder, characterIndex, 1)
        If characterValue = " " Or characterValue = vbTab Or _
           characterValue = "\" Then Exit For
        identifier = identifier & characterValue
    Next characterIndex
    VTSequenceIdentifierFromFieldCode = identifier
End Function

Public Function VTEquationSequenceResultText( _
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
    ByVal sequenceOrdinal As Long, _
    Optional ByVal restartLevel As Long = 0) As String

    If sequenceOrdinal < 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The Equation sequence ordinal must be positive."
    End If

    ' Keep every caption as a real native Word SEQ field. Chapter/section
    ' boundaries are resolved by VisualTeX from the same OutlineLevel/ListString
    ' scan that produces the visible heading prefix. The first Equation in each
    ' resolved scope explicitly restarts the native Equation sequence at 1;
    ' later Equations remain ordinary flowing SEQ fields. This avoids relying on
    ' Word's \s Heading-style lookup, which does not match custom outline-level
    ' headings, while retaining Word's native caption/cross-reference machinery.
    ' Helpers are external to OMath, so use the same native \* ARABIC switch as
    ' Windows production; it is also required for Word's own caption target span.
    VTEquationSequenceFieldCodeForOrdinal = _
        " SEQ " & VTEquationSequenceFieldText(equationLabelName)
    If restartLevel = 1 Or restartLevel = 2 Then
        If sequenceOrdinal = 1 Then
            VTEquationSequenceFieldCodeForOrdinal = _
                VTEquationSequenceFieldCodeForOrdinal & " \r 1"
        End If
    ElseIf restartLevel <> 0 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The Equation sequence restart level is invalid."
    End If
    VTEquationSequenceFieldCodeForOrdinal = _
        VTEquationSequenceFieldCodeForOrdinal & " \* ARABIC "
End Function

Public Function VTNormalizeEquationFieldCode( _
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

Public Function VTEquationSequenceFieldHasOrdinal( _
    ByVal sequenceField As Field, _
    ByVal equationLabelName As String, _
    ByVal sequenceOrdinal As Long, _
    Optional ByVal restartLevel As Long = 0) As Boolean

    Dim actualCode As String
    Dim expectedCode As String

    If sequenceField Is Nothing Or sequenceOrdinal < 1 Then Exit Function
    actualCode = VTNormalizeEquationFieldCode(sequenceField.Code.Text)
    expectedCode = VTNormalizeEquationFieldCode( _
        VTEquationSequenceFieldCodeForOrdinal( _
            equationLabelName, sequenceOrdinal, restartLevel))
    VTEquationSequenceFieldHasOrdinal = _
        StrComp(actualCode, expectedCode, vbTextCompare) = 0 And _
        VTEquationSequenceResultText(sequenceField) = CStr(sequenceOrdinal)
End Function

Public Function VTEquationSequenceFieldHasSwitch( _
    ByVal sequenceField As Field, _
    ByVal switchName As String) As Boolean

    Dim paddedCode As String

    If sequenceField Is Nothing Or Len(Trim$(switchName)) = 0 Then
        Exit Function
    End If
    paddedCode = " " & LCase$(VTNormalizeEquationFieldCode( _
        sequenceField.Code.Text)) & " "
    VTEquationSequenceFieldHasSwitch = _
        InStr(1, paddedCode, _
            " \" & LCase$(Trim$(switchName)) & " ", _
            vbBinaryCompare) > 0
End Function

Private Function VTEquationSequenceFieldRestartValue( _
    ByVal sequenceField As Field) As Long

    Dim paddedCode As String
    Dim remainder As String
    Dim valueToken As String
    Dim switchStart As Long
    Dim tokenEnd As Long
    Dim restartValue As Double

    VTEquationSequenceFieldRestartValue = -1
    If sequenceField Is Nothing Then Exit Function
    paddedCode = " " & LCase$(VTNormalizeEquationFieldCode( _
        sequenceField.Code.Text)) & " "
    switchStart = InStr(1, paddedCode, " \r ", vbBinaryCompare)
    If switchStart = 0 Then Exit Function
    remainder = LTrim$(Mid$(paddedCode, switchStart + 4))
    tokenEnd = InStr(1, remainder, " ", vbBinaryCompare)
    If tokenEnd > 0 Then
        valueToken = Left$(remainder, tokenEnd - 1)
    Else
        valueToken = remainder
    End If
    If Len(valueToken) = 0 Or Not IsNumeric(valueToken) Then Exit Function
    restartValue = CDbl(valueToken)
    If restartValue < 0 Or restartValue > 2147483647# Or _
       restartValue <> Fix(restartValue) Then Exit Function
    VTEquationSequenceFieldRestartValue = CLng(restartValue)
End Function

Public Function VTEquationSequenceFieldIsCompatible( _
    ByVal sequenceField As Field, _
    ByVal sequenceName As String, _
    ByVal hiddenIncrement As Boolean, _
    ByVal currentValue As Boolean, _
    ByVal restartAtOne As Boolean) As Boolean

    Dim hasHidden As Boolean
    Dim hasCurrent As Boolean
    Dim hasRestartLevel As Boolean
    Dim hasExplicitRestart As Boolean

    If sequenceField Is Nothing Or sequenceField.Type <> wdFieldSequence Then _
        Exit Function
    If StrComp(VTSequenceIdentifierFromFieldCode(sequenceField.Code.Text), _
       sequenceName, vbTextCompare) <> 0 Then Exit Function

    hasHidden = VTEquationSequenceFieldHasSwitch(sequenceField, "h")
    hasCurrent = VTEquationSequenceFieldHasSwitch(sequenceField, "c")
    hasRestartLevel = VTEquationSequenceFieldHasSwitch(sequenceField, "s")
    hasExplicitRestart = VTEquationSequenceFieldHasSwitch(sequenceField, "r")
    If hasHidden <> hiddenIncrement Or hasCurrent <> currentValue Then _
        Exit Function
    If hasRestartLevel Then Exit Function

    If currentValue Then
        If hasExplicitRestart Then Exit Function
    ElseIf restartAtOne Then
        If Not hasExplicitRestart Or _
           VTEquationSequenceFieldRestartValue(sequenceField) <> 1 Then _
            Exit Function
    ElseIf hasExplicitRestart Then
        Exit Function
    End If

    ' Word is free to add harmless formatting switches such as MERGEFORMAT.
    ' Validate only the switches that change sequence semantics so a healthy
    ' field is not rebuilt merely because Word normalized its presentation.
    VTEquationSequenceFieldIsCompatible = True
End Function

Public Sub VTApplyEquationSequenceOrdinal( _
    ByVal sequenceField As Field, _
    ByVal equationLabelName As String, _
    ByVal sequenceOrdinal As Long, _
    Optional ByVal restartLevel As Long = 0)

    Dim expectedCode As String

    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The Equation SEQ field is missing."
    End If
    expectedCode = VTEquationSequenceFieldCodeForOrdinal( _
        equationLabelName, sequenceOrdinal, restartLevel)

    ' Never replace the Field object: changing only its code keeps the native
    ' Equation caption identity in place. In chapter/section mode only formulas
    ' that gain or lose first-in-scope status need a code change; all other
    ' renumbering remains a normal flowing SEQ update.
    If StrComp(VTNormalizeEquationFieldCode(sequenceField.Code.Text), _
       VTNormalizeEquationFieldCode(expectedCode), vbTextCompare) <> 0 Then
        sequenceField.Code.Text = expectedCode
    End If
    sequenceField.Update

    If Not VTEquationSequenceFieldHasOrdinal( _
       sequenceField, equationLabelName, sequenceOrdinal, restartLevel) Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "Word did not update the flowing native Equation SEQ field" & _
            " [code=" & sequenceField.Code.Text & _
            "; result=" & sequenceField.Result.Text & _
            "; expected=" & CStr(sequenceOrdinal) & "]."
    End If
End Sub

Public Function VTRefreshFormattedSequenceBookmark( _
    ByVal documentObject As Document, _
    ByVal sequenceField As Field, _
    ByVal formulaId As String, _
    Optional ByVal headingPrefixOverride As String = "", _
    Optional ByVal headingPrefixResolved As Boolean = False) As Range

    Dim formulaRange As Range
    Dim helperParagraph As Range
    Dim formulaShape As InlineShape
    Dim formulaContainerRange As Range
    Dim prefixRange As Range
    Dim suffixRange As Range
    Dim bookmarkRange As Range
    Dim numberRange As Range
    Dim captionRange As Range
    Dim numberingMode As String
    Dim headingPrefix As String
    Dim prefixText As String
    Dim sequenceBookmarkName As String
    Dim helperStart As Long
    Dim fieldStart As Long
    Dim numberFontName As String
    Dim numberFontSize As Single
    Dim fieldEnd As Long

    If documentObject Is Nothing Or sequenceField Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7596, "VisualTeX", _
            "The formatted Equation sequence target is invalid."
    End If
    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7596, "VisualTeX", _
            "The formatted Equation sequence has no visible formula."
    End If
    numberingMode = VTEquationNumberingMode(documentObject)
    If numberingMode <> VT_WORD_NUMBERING_MODE_SEQUENCE Then
        If headingPrefixResolved Then
            headingPrefix = headingPrefixOverride
        Else
            headingPrefix = VTEquationHeadingPrefix( _
                formulaRange, numberingMode)
        End If
        prefixText = headingPrefix & _
            VTEquationNumberingDisplaySeparator(documentObject)
    End If

    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate

    ' Numbered image formulas own their real SEQ directly in the visible
    ' center/right-tab row. Normalize only the static prefix/parentheses around
    ' that one field; there is deliberately no external helper paragraph.
    If formulaRange.InlineShapes.Count = 1 And _
       formulaRange.OMaths.Count = 0 And _
       helperParagraph.Start = formulaRange.Paragraphs(1).Range.Start Then
        Set formulaShape = formulaRange.InlineShapes(1)
        Set formulaContainerRange = _
            VTVisualTeXImageContainerRange(formulaShape)
        helperStart = helperParagraph.Start
        fieldStart = VTEquationFieldStart(sequenceField)
        fieldEnd = VTEquationFieldEnd(sequenceField)
        sequenceBookmarkName = _
            VTEquationSequenceNumberBookmarkName(formulaId)
        If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
            documentObject.Bookmarks(sequenceBookmarkName).Delete
        End If

        Set prefixRange = documentObject.Range( _
            Start:=formulaContainerRange.End, End:=fieldStart)
        prefixRange.Text = vbTab & "(" & prefixText
        VTUnbiasNumberSeparators documentObject, prefixRange.Start, _
            prefixRange.Start + Len(vbTab & "(" & prefixText)
        Set helperParagraph = documentObject.Range( _
            Start:=helperStart, End:=helperStart).Paragraphs(1).Range.Duplicate
        Set sequenceField = VTFindEquationSequenceField(helperParagraph)
        If sequenceField Is Nothing Then
            Err.Raise vbObjectError + 7596, "VisualTeX", _
                "The direct image Equation lost its SEQ while formatting the prefix."
        End If
        fieldEnd = VTEquationFieldEnd(sequenceField)
        Set suffixRange = documentObject.Range( _
            Start:=fieldEnd, End:=helperParagraph.End - 1)
        suffixRange.Text = ")"

        Set helperParagraph = documentObject.Range( _
            Start:=helperStart, End:=helperStart).Paragraphs(1).Range.Duplicate
        Set formulaShape = helperParagraph.InlineShapes(1)
        Set formulaContainerRange = _
            VTVisualTeXImageContainerRange(formulaShape)
        Set sequenceField = VTFindEquationSequenceField(helperParagraph)
        Set numberRange = documentObject.Range( _
            Start:=formulaContainerRange.End + 1, _
            End:=helperParagraph.End - 1)
        Set bookmarkRange = documentObject.Range( _
            Start:=numberRange.Start + 1, _
            End:=sequenceField.Result.End)
        documentObject.Bookmarks.Add _
            name:=sequenceBookmarkName, Range:=bookmarkRange
        VTSetEquationNumberBookmarkExact _
            documentObject, formulaId, numberRange
        Set captionRange = numberRange.Duplicate
        captionRange.Collapse wdCollapseEnd
        If documentObject.Bookmarks.Exists( _
           VTEquationCaptionBookmarkName(formulaId)) Then
            documentObject.Bookmarks( _
                VTEquationCaptionBookmarkName(formulaId)).Delete
        End If
        documentObject.Bookmarks.Add _
            name:=VTEquationCaptionBookmarkName(formulaId), _
            Range:=captionRange
        VTImageFormulaNumberFormat formulaShape, numberFontName, numberFontSize
        VTApplyStaticImageEquationNumberFormatting _
            numberRange, numberFontSize, numberFontName
        Set VTRefreshFormattedSequenceBookmark = bookmarkRange.Duplicate
        Exit Function
    End If

    If Not VTHelperParagraphOwnsNativeEquationSequence( _
       helperParagraph) Then
        Err.Raise vbObjectError + 7596, "VisualTeX", _
            "The formatted Equation sequence helper is invalid."
    End If
    helperStart = helperParagraph.Start
    fieldStart = VTEquationFieldStart(sequenceField)
    fieldEnd = VTEquationFieldEnd(sequenceField)
    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        documentObject.Bookmarks(sequenceBookmarkName).Delete
    End If

    ' Remove only the static formatting text around the one live SEQ field.
    ' The field object and its native Word sequence identity stay in place.
    Set suffixRange = documentObject.Range( _
        Start:=fieldEnd, End:=helperParagraph.End - 1)
    If suffixRange.End > suffixRange.Start Then suffixRange.Delete
    Set helperParagraph = documentObject.Range( _
        Start:=helperStart, End:=helperStart).Paragraphs(1).Range.Duplicate
    Set sequenceField = helperParagraph.Fields(1)
    fieldStart = VTEquationFieldStart(sequenceField)
    Set prefixRange = documentObject.Range( _
        Start:=helperParagraph.Start, End:=fieldStart)
    If prefixRange.End > prefixRange.Start Then prefixRange.Delete

    Set helperParagraph = documentObject.Range( _
        Start:=helperStart, End:=helperStart).Paragraphs(1).Range.Duplicate
    Set sequenceField = helperParagraph.Fields(1)
    If Len(prefixText) > 0 Then
        Set prefixRange = documentObject.Range( _
            Start:=helperParagraph.Start, End:=helperParagraph.Start)
        prefixRange.Text = prefixText
        VTUnbiasNumberSeparators documentObject, _
            helperParagraph.Start, helperParagraph.Start + Len(prefixText)
        Set helperParagraph = documentObject.Range( _
            Start:=helperStart, End:=helperStart).Paragraphs(1).Range.Duplicate
        Set sequenceField = helperParagraph.Fields(1)
        Set bookmarkRange = documentObject.Range( _
            Start:=helperParagraph.Start, _
            End:=sequenceField.Result.End)
    Else
        Set bookmarkRange = sequenceField.Result.Duplicate
    End If
    documentObject.Bookmarks.Add _
        name:=sequenceBookmarkName, Range:=bookmarkRange
    VTFormatHiddenEquationParagraph helperParagraph
    Set VTRefreshFormattedSequenceBookmark = bookmarkRange.Duplicate
End Function

Public Function VTEquationNumberTextForFormula( _
    ByVal documentObject As Document, _
    ByVal formulaId As String) As String

    Dim sequenceBookmarkName As String
    Dim numberText As String
    Dim nativeBookmarkName As String
    Dim nativeMath As OMath
    Dim numberField As Field
    Dim numberRange As Range

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then
        Exit Function
    End If
    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        numberText = documentObject.Bookmarks( _
            sequenceBookmarkName).Range.Text
    Else
        ' Current managed OMML owns its SEQ directly inside the equation array;
        ' it intentionally has no legacy VT_N_ helper. Resolve the number from
        ' the retained edit identity when a complete managed cache is present.
        nativeBookmarkName = VTNativeFormulaBookmarkName(formulaId)
        If Not documentObject.Bookmarks.Exists(nativeBookmarkName) Then _
            Exit Function
        Set nativeMath = VTNativeMathForBookmark( _
            documentObject.Bookmarks(nativeBookmarkName))
        If nativeMath Is Nothing Then Exit Function
        Set numberField = VTPureNativeNumberField(nativeMath.Range)
        If numberField Is Nothing Then Exit Function
        Set numberRange = VTPureNativeVisibleNumberRange( _
            nativeMath.Range, numberField)
        If numberRange Is Nothing Then Exit Function
        numberText = numberRange.Text
    End If
    numberText = Replace$(numberText, vbCr, "")
    numberText = Replace$(numberText, Chr$(7), "")
    numberText = Replace$(numberText, Chr$(11), "")
    VTEquationNumberTextForFormula = Trim$(numberText)
End Function

Public Function VTEquationSequenceOrdinal( _
    ByVal documentObject As Document, _
    ByVal targetField As Field, _
    ByVal equationLabelName As String) As Long

    Dim candidate As Field
    Dim sequenceOrdinal As Long
    Dim targetStart As Long
    Dim candidateStart As Long

    If documentObject Is Nothing Or targetField Is Nothing Then Exit Function
    targetStart = VTEquationFieldStart(targetField)

    ' Prefer Word's own evaluated result. Positional field counting cannot model
    ' legal native Equation captions that use \r, \c or \s switches.
    targetField.Update
    sequenceOrdinal = VTFirstPositiveIntegerInText( _
        VTEquationSequenceResultText(targetField))
    If sequenceOrdinal > 0 Then
        VTEquationSequenceOrdinal = sequenceOrdinal
        Exit Function
    End If

    ' Compatibility fallback for a Word build that leaves a fresh result empty:
    ' count every recognized same-sequence field at or before the target by its
    ' physical field position. Never depend on Fields collection enumeration.
    For Each candidate In documentObject.Fields
        If VTIsNativeEquationSequenceField(candidate, equationLabelName) Then
            candidateStart = VTEquationFieldStart(candidate)
            If candidateStart <= targetStart Then _
                sequenceOrdinal = sequenceOrdinal + 1
        End If
    Next candidate
    VTEquationSequenceOrdinal = sequenceOrdinal
End Function

Private Function VTHasNativeEquationSequenceAfter( _
    ByVal documentObject As Document, _
    ByVal sequenceField As Field) As Boolean

    Dim candidate As Field
    Dim trailingRange As Range
    Dim trailingStart As Long

    If documentObject Is Nothing Or sequenceField Is Nothing Then
        VTHasNativeEquationSequenceAfter = True
        Exit Function
    End If
    trailingStart = sequenceField.Result.Paragraphs(1).Range.End
    If trailingStart >= documentObject.Content.End Then Exit Function

    ' Sequential insertion normally leaves the new helper at the document tail.
    ' Restrict the probe to content after that helper: an append therefore checks
    ' an empty/tiny Range regardless of how many numbered formulas precede it.
    Set trailingRange = documentObject.Range( _
        Start:=trailingStart, End:=documentObject.Content.End)
    For Each candidate In trailingRange.Fields
        If VTIsNativeEquationSequenceField( _
           candidate, VTNativeEquationLabelName()) Then
            VTHasNativeEquationSequenceAfter = True
            Exit Function
        End If
    Next candidate
End Function

Private Function VTHasManagedEquationNumberAfter( _
    ByVal documentObject As Document, _
    ByVal sequenceField As Field) As Boolean

    Dim candidateBookmark As Bookmark
    Dim trailingRange As Range
    Dim trailingStart As Long
    Dim bookmarkName As String

    If documentObject Is Nothing Or sequenceField Is Nothing Then
        VTHasManagedEquationNumberAfter = True
        Exit Function
    End If
    trailingStart = sequenceField.Result.Paragraphs(1).Range.End
    If trailingStart >= documentObject.Content.End Then Exit Function

    ' A tail fast path is valid only when this helper is physically after every
    ' later VisualTeX numbered formula as well as every later Equation SEQ. The
    ' VT_R_ visible-number Bookmark catches a managed formula even if Word's live
    ' Fields collection is temporarily stale while an empty middle paragraph is
    ' being converted into a new numbered OMath.
    Set trailingRange = documentObject.Range( _
        Start:=trailingStart, End:=documentObject.Content.End)
    For Each candidateBookmark In trailingRange.Bookmarks
        bookmarkName = candidateBookmark.Name
        If Left$(bookmarkName, Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_NUMBER_BOOKMARK_PREFIX Or _
           Left$(bookmarkName, _
               Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)) = _
               VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX Then
            VTHasManagedEquationNumberAfter = True
            Exit Function
        End If
    Next candidateBookmark
End Function

Private Function VTCanUseEquationTailFastPath( _
    ByVal documentObject As Document, _
    ByVal sequenceField As Field) As Boolean

    If documentObject Is Nothing Or sequenceField Is Nothing Then Exit Function
    If VTHasNativeEquationSequenceAfter( _
       documentObject, sequenceField) Then
        VTWordPerformanceMark "equation-tail-fast-blocked-native-after"
        Exit Function
    End If
    If VTHasManagedEquationNumberAfter( _
       documentObject, sequenceField) Then
        VTWordPerformanceMark "equation-tail-fast-blocked-managed-after"
        Exit Function
    End If
    VTCanUseEquationTailFastPath = True
End Function

Private Function VTHasOtherManagedEquationNumber( _
    ByVal documentObject As Document, _
    ByVal formulaId As String) As Boolean

    Dim candidateBookmark As Bookmark
    Dim bookmarkName As String
    Dim ownNumberBookmarkName As String
    Dim ownSequenceBookmarkName As String

    If documentObject Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then
        VTHasOtherManagedEquationNumber = True
        Exit Function
    End If
    ownNumberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    ownSequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    For Each candidateBookmark In documentObject.Bookmarks
        bookmarkName = candidateBookmark.Name
        If Left$(bookmarkName, _
           Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_NUMBER_BOOKMARK_PREFIX Then
            If StrComp(bookmarkName, ownNumberBookmarkName, _
               vbTextCompare) <> 0 Then
                VTHasOtherManagedEquationNumber = True
                Exit Function
            End If
        ElseIf Left$(bookmarkName, _
           Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX Then
            If StrComp(bookmarkName, ownSequenceBookmarkName, _
               vbTextCompare) <> 0 Then
                VTHasOtherManagedEquationNumber = True
                Exit Function
            End If
        End If
    Next candidateBookmark
End Function

Public Function VTEquationFieldStart( _
    ByVal sequenceField As Field) As Long

    If sequenceField Is Nothing Or sequenceField.Code.Start <= 0 Then
        Err.Raise vbObjectError + 7534, "VisualTeX", _
            "The Equation field start boundary is invalid."
    End If
    VTEquationFieldStart = sequenceField.Code.Start - 1
End Function

Public Function VTEquationFieldEnd( _
    ByVal sequenceField As Field) As Long

    If sequenceField Is Nothing Or sequenceField.Result.End < 0 Then
        Err.Raise vbObjectError + 7535, "VisualTeX", _
            "The Equation field end boundary is invalid."
    End If
    If VTIsImagePlaceRef(sequenceField) Then
        ' A MACROBUTTON label lives in Code and has no separate result story.
        VTEquationFieldEnd = sequenceField.Code.End + 1
    Else
        VTEquationFieldEnd = sequenceField.Result.End + 1
    End If
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

Public Function VTEquationCaptionBookmarkName( _
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

Public Sub VTSetCollapsedEquationCaptionBookmark( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal paragraphRange As Range)

    Dim bookmarkName As String
    Dim captionRange As Range

    If documentObject Is Nothing Or paragraphRange Is Nothing Then
        Err.Raise vbObjectError + 7567, "VisualTeX", _
            "The Equation caption identity target is missing."
    End If

    ' The historical procedure name is retained while schema-7 documents migrate,
    ' but production VT_C_ now matches Windows: it covers the complete native SEQ
    ' helper paragraph, including its paragraph mark. A collapsed tail Bookmark
    ' leaves Word for Mac without a real caption span, so InsertCrossReference can
    ' create its private _Ref at the preceding formula paragraph instead.
    Set captionRange = paragraphRange.Paragraphs(1).Range.Duplicate
    bookmarkName = VTEquationCaptionBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        documentObject.Bookmarks(bookmarkName).Delete
    End If
    documentObject.Bookmarks.Add _
        name:=bookmarkName, Range:=captionRange
End Sub

Public Function VTEquationCaptionBookmarkIsCollapsedInParagraph( _
    ByVal captionRange As Range, _
    ByVal paragraphRange As Range) As Boolean

    Dim exactParagraph As Range

    If captionRange Is Nothing Or paragraphRange Is Nothing Then Exit Function
    Set exactParagraph = paragraphRange.Paragraphs(1).Range.Duplicate
    If captionRange.Start <> exactParagraph.Start Or _
       captionRange.End <> exactParagraph.End Then Exit Function
    If captionRange.Information(wdWithInTable) <> _
       exactParagraph.Information(wdWithInTable) Then Exit Function
    VTEquationCaptionBookmarkIsCollapsedInParagraph = True
End Function

Public Function VTEquationNumberBookmarkName( _
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

Public Function VTEquationSequenceNumberBookmarkName( _
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

Public Function VTEquationSequenceFieldForBookmark( _
    ByVal documentObject As Document, _
    ByVal sequenceBookmarkName As String) As Field

    Dim candidate As Field
    Dim bookmarkRange As Range
    Dim probeRange As Range
    Dim helperRange As Range
    Dim match As Field
    Dim matchCount As Long
    Dim equationLabelName As String
    Dim probeStart As Long
    Dim probeEnd As Long

    If documentObject Is Nothing Or Len(sequenceBookmarkName) = 0 Then
        Exit Function
    End If
    If Not documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        Exit Function
    End If
    Set bookmarkRange = documentObject.Bookmarks( _
        sequenceBookmarkName).Range.Duplicate
    equationLabelName = VTNativeEquationLabelName()

    ' Healthy VisualTeX numbering keeps VT_N_ on the SEQ result (sequence mode)
    ' or on the short prefix + SEQ result inside the dedicated helper paragraph
    ' (chapter/section modes). Resolve that field locally instead of enumerating
    ' every Field in the document for every formula. The former global lookup
    ' turned an otherwise linear reconciliation into O(n^2) after a few dozen
    ' numbered formulas.
    probeStart = bookmarkRange.Start - 2
    If probeStart < 0 Then probeStart = 0
    probeEnd = bookmarkRange.End + 2
    If probeEnd > documentObject.Content.End Then
        probeEnd = documentObject.Content.End
    End If
    Set probeRange = documentObject.Range( _
        Start:=probeStart, End:=probeEnd)
    For Each candidate In probeRange.Fields
        If VTIsNativeEquationSequenceField(candidate, equationLabelName) Then
            If (candidate.Result.Start <= bookmarkRange.Start And _
                candidate.Result.End >= bookmarkRange.End) Or _
               (bookmarkRange.Start <= candidate.Result.Start And _
                bookmarkRange.End >= candidate.Result.End) Or _
               (candidate.Result.Start >= bookmarkRange.Start And _
                candidate.Result.End <= bookmarkRange.End) Then
                matchCount = matchCount + 1
                Set match = candidate
            End If
        End If
    Next candidate
    If matchCount = 1 Then
        Set VTEquationSequenceFieldForBookmark = match
        Exit Function
    ElseIf matchCount > 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The Equation sequence Bookmark resolves to multiple local fields."
    End If

    ' A field-result Bookmark can briefly collapse while Word updates the field.
    ' Its helper paragraph is still a tiny, formula-local recovery surface.
    Set helperRange = bookmarkRange.Paragraphs(1).Range.Duplicate
    For Each candidate In helperRange.Fields
        If VTIsNativeEquationSequenceField(candidate, equationLabelName) Then
            matchCount = matchCount + 1
            Set match = candidate
        End If
    Next candidate
    If matchCount = 1 Then
        Set VTEquationSequenceFieldForBookmark = match
        Exit Function
    ElseIf matchCount > 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The Equation sequence helper contains multiple sequence fields."
    End If

    ' Keep one compatibility fallback for malformed legacy documents. Current
    ' documents should never reach it, so correctness is preserved without
    ' putting a document-wide Fields scan on every normal insert/edit/update.
    For Each candidate In documentObject.Fields
        If VTIsNativeEquationSequenceField(candidate, equationLabelName) Then
            If (candidate.Result.Start <= bookmarkRange.Start And _
                candidate.Result.End >= bookmarkRange.End) Or _
               (bookmarkRange.Start <= candidate.Result.Start And _
                bookmarkRange.End >= candidate.Result.End) Then
                matchCount = matchCount + 1
                Set match = candidate
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

Private Sub VTEnsureNativeEquationHelperFrame( _
    ByVal paragraphRange As Range)

    Dim documentObject As Document
    Dim helperFrame As Frame
    Dim frameRange As Range
    Dim frameHorizontalPosition As Single
    Dim frameVerticalPosition As Single
    Dim pageWidth As Single
    Dim pageHeight As Single

    If paragraphRange Is Nothing Then Exit Sub
    If paragraphRange.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The native Equation helper Frame cannot live inside a table."
    End If

    Set documentObject = paragraphRange.Document
    Set frameRange = paragraphRange.Duplicate
    If frameRange.Frames.Count > 1 Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The native Equation helper belongs to multiple Word Frames."
    ElseIf frameRange.Frames.Count = 1 Then
        Set helperFrame = frameRange.Frames(1)
    Else
        Set helperFrame = documentObject.Frames.Add(Range:=frameRange)
    End If
    If frameRange.Sections.Count < 1 Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The native Equation helper has no page geometry."
    End If
    pageWidth = frameRange.Sections(1).PageSetup.PageWidth
    pageHeight = frameRange.Sections(1).PageSetup.PageHeight
    frameHorizontalPosition = 0!
    frameVerticalPosition = 0!
    If pageWidth > 0.1 Then _
        frameHorizontalPosition = pageWidth - 0.1
    If pageHeight > 0.1 Then _
        frameVerticalPosition = pageHeight - 0.1

    ' Match the mature Windows production geometry exactly. Word clamps negative
    ' coordinates to the top-left, exposing the Frame edge and allowing the first
    ' private _Ref to collapse onto the formula row. A numeric 0.1-point square
    ' at the bottom-right page boundary keeps the native SEQ in the main story
    ' while clipping its rendered content wholly beyond the page.
    With helperFrame
        .HeightRule = wdFrameExact
        .Height = 0.1
        .WidthRule = wdFrameExact
        .Width = 0.1
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .HorizontalPosition = frameHorizontalPosition
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .VerticalPosition = frameVerticalPosition
        .TextWrap = False
        .LockAnchor = True
    End With
    paragraphRange.Borders.Enable = False
End Sub

Public Sub VTFormatHiddenEquationParagraph(ByVal paragraphRange As Range)
    Dim candidateField As Field
    Dim equationLabelName As String
    Dim visibleSize As Single

    If paragraphRange Is Nothing Then Exit Sub
    visibleSize = VTVisibleEquationNumberFontSize(paragraphRange.Document)
    equationLabelName = VTNativeEquationLabelName()

    ' Word's native cross-reference inventory expects a native caption paragraph.
    ' Some locked/custom documents reject the built-in style, so retain Windows'
    ' best-effort behavior and enforce every production property explicitly.
    On Error Resume Next
    paragraphRange.Style = wdStyleCaption
    paragraphRange.ListFormat.RemoveNumbers wdNumberParagraph
    Err.Clear
    On Error GoTo 0

    ' F9 rebuilds a field result from the formatting around the field, discarding
    ' direct formatting previously applied only to candidateField.Result. Keep
    ' the complete helper paragraph in the production number face so a rebuilt
    ' SEQ and every visible REF that copies it remain Cambria Math Position 0.
    ' A page-boundary 0.1-point native Frame clips that full-size source without
    ' changing the character formatting Word copies into a visible REF.
    With paragraphRange.Font
        .Size = visibleSize
        .Hidden = False
        .Color = wdColorAutomatic
        .Name = VT_WORD_EQUATION_NUMBER_FONT_NAME
        .NameAscii = VT_WORD_EQUATION_NUMBER_FONT_NAME
        .NameOther = VT_WORD_EQUATION_NUMBER_FONT_NAME
        .Position = 0
    End With
    With paragraphRange.ParagraphFormat
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
        .WidowControl = False
    End With

    ' Reinforce the current SEQ result as well as its surrounding field/paragraph
    ' source formatting. This covers the first native cross-reference insertion
    ' before any F9 cycle and keeps the invariant explicit for integrity checks.
    For Each candidateField In paragraphRange.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, equationLabelName) Then
            With candidateField.Result.Font
                .Size = visibleSize
                .Hidden = False
                .Color = wdColorAutomatic
                .Name = VT_WORD_EQUATION_NUMBER_FONT_NAME
                .NameAscii = VT_WORD_EQUATION_NUMBER_FONT_NAME
                .NameOther = VT_WORD_EQUATION_NUMBER_FONT_NAME
                .Position = 0
            End With
        End If
    Next candidateField

    VTEnsureNativeEquationHelperFrame paragraphRange
End Sub

Public Sub VTSetEquationNumberBookmarkExact( _
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

Public Function VTNativeEquationArrayMarkerRange( _
    ByVal formulaRange As Range) As Range

    Dim markerRange As Range
    Dim rootMath As OMath
    Dim arrayFunction As OMathFunction
    Dim slotFunction As OMathFunction
    Dim numberDelimiter As OMathFunction

    If formulaRange Is Nothing Then Exit Function
    If formulaRange.OMaths.Count <> 1 Then Exit Function
    Set rootMath = VTNativeMathRoot(formulaRange.OMaths(1))
    ' A native number is body#(number) in the first row of an equation array,
    ' so its # is the character before that row's last delimiter. Reading the
    ' structure costs a fraction of a Find, which takes about 9 ms an equation
    ' on Word for Mac and several times more once the session has shown an
    ' image formula. Any other shape keeps the Find for the last #.
    For Each arrayFunction In rootMath.Functions
        If arrayFunction.Type = wdOMathFunctionEqArray Then
            For Each slotFunction In arrayFunction.EqArray.E(1).Functions
                If slotFunction.Type = wdOMathFunctionDelim Then _
                    Set numberDelimiter = slotFunction
            Next slotFunction
            Exit For
        End If
    Next arrayFunction
    If Not numberDelimiter Is Nothing Then
        If numberDelimiter.Range.Start > rootMath.Range.Start Then
            Set markerRange = rootMath.Range.Document.Range( _
                numberDelimiter.Range.Start - 1, numberDelimiter.Range.Start)
            If markerRange.Text = "#" Then
                Set VTNativeEquationArrayMarkerRange = markerRange
                Exit Function
            End If
        End If
    End If
    Set markerRange = rootMath.Range.Duplicate
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

Public Function VTNativeEquationArrayPlaceholderRange( _
    ByVal formulaRange As Range, _
    Optional ByVal expectedToken As String = VT_WORD_NATIVE_NUMBER_CACHE_PLACEHOLDER) As Range

    Dim exactRange As Range
    Dim placeholderRange As Range
    Dim shellXml As String
    Dim arrayXmlStart As Long
    Dim markerXmlStart As Long
    Dim markerTextStart As Long
    Dim markerTextEnd As Long
    Dim placeholderXmlStart As Long
    Dim delimiterXmlStart As Long
    Dim delimiterXmlEnd As Long
    Dim arrayXmlEnd As Long
    Dim delimiterXml As String
    Dim shellTailXml As String
    Dim arrayFunction As OMathFunction
    Dim slotFunction As OMathFunction
    Dim arrayMath As OMath
    Dim slotMath As OMath

    If formulaRange Is Nothing Then Exit Function
    If formulaRange.OMaths.Count <> 1 Then Exit Function
    Set exactRange = VTNativeMathRoot(formulaRange.OMaths(1)).Range.Duplicate
    shellXml = exactRange.WordOpenXML
    arrayXmlStart = InStr(1, shellXml, "<m:eqArr", vbBinaryCompare)
    If arrayXmlStart = 0 Then Exit Function
    ' Word splits the sentinel across math runs when letters and digits inherit
    ' different bold/italic properties. Validate the last delimiter's combined
    ' text structurally; never search for the literal token in serialized XML.
    delimiterXmlEnd = InStrRev(shellXml, "</m:d>", -1, vbBinaryCompare)
    If delimiterXmlEnd <= arrayXmlStart Then Exit Function
    delimiterXmlStart = VTXmlElementStartBefore( _
        shellXml, "m:d", delimiterXmlEnd)
    If delimiterXmlStart <= arrayXmlStart Then Exit Function
    markerXmlStart = InStrRev( _
        Left$(shellXml, delimiterXmlStart - 1), "#", -1, vbBinaryCompare)
    If markerXmlStart = 0 Then Exit Function
    markerTextStart = InStrRev( _
        Left$(shellXml, markerXmlStart), "<m:t", -1, vbBinaryCompare)
    markerTextEnd = InStr( _
        markerXmlStart, shellXml, "</m:t>", vbBinaryCompare)
    If markerTextStart < arrayXmlStart Or markerTextEnd = 0 Or _
       markerTextEnd >= delimiterXmlStart Then Exit Function
    arrayXmlEnd = InStr( _
        delimiterXmlEnd, shellXml, "</m:eqArr>", vbBinaryCompare)
    If delimiterXmlStart = 0 Or delimiterXmlEnd = 0 Or _
       arrayXmlEnd = 0 Then Exit Function
    If InStr(markerXmlStart, Left$(shellXml, arrayXmlEnd), _
       "w:fldChar", vbBinaryCompare) > 0 Then Exit Function
    delimiterXml = Mid$(shellXml, delimiterXmlStart, _
        delimiterXmlEnd + Len("</m:d>") - delimiterXmlStart)
    If VTXmlTextContent(delimiterXml) <> _
       expectedToken Then Exit Function
    shellTailXml = Mid$(shellXml, _
        delimiterXmlEnd + Len("</m:d>"), _
        arrayXmlEnd - delimiterXmlEnd - Len("</m:d>"))
    shellTailXml = Replace$(shellTailXml, "</m:e>", "")
    shellTailXml = Replace$(shellTailXml, vbCr, "")
    shellTailXml = Replace$(shellTailXml, vbLf, "")
    shellTailXml = Replace$(shellTailXml, vbTab, "")
    shellTailXml = Replace$(shellTailXml, " ", "")
    If Len(shellTailXml) <> 0 Then Exit Function
    ' The slot is already strictly owned by the XML checks above. Use Word's
    ' delimiter argument Range, whose length includes surrogate pairs correctly.
    ' Find("0") cannot match a bold mathematical digit inherited from a user's
    ' paragraph style, and expanding by ASCII length truncates bold letters.
    On Error Resume Next
    For Each arrayFunction In exactRange.OMaths(1).Functions
        If arrayFunction.Type = wdOMathFunctionEqArray Then
            Set arrayMath = arrayFunction.EqArray.E(1)
            For Each slotFunction In arrayMath.Functions
                If slotFunction.Type = wdOMathFunctionDelim Then
                    If slotFunction.Range.Start > _
                       VTNativeEquationArrayMarkerRange(exactRange).Start Then
                        Set slotMath = slotFunction.Delim.E(1)
                        Set VTNativeEquationArrayPlaceholderRange = _
                            slotMath.Range.Duplicate
                        Exit Function
                    End If
                End If
            Next slotFunction
        End If
    Next arrayFunction
    Err.Clear
    On Error GoTo 0
    Set placeholderRange = exactRange.Duplicate
    With placeholderRange.Find
        .ClearFormatting
        ' Word exposes math letters through Range.Text as Unicode mathematical
        ' alphabet characters even though WordOpenXML retains the ASCII sentinel.
        ' The strict delimiter XML check above owns the slot; anchor its final
        ' digit and expand by the known sentinel length in the document Range.
        .Text = Right$(expectedToken, 1)
        .Forward = False
        .Wrap = wdFindStop
        .Format = False
    End With
    If Not placeholderRange.Find.Execute Then Exit Function
    If placeholderRange.End - Len(expectedToken) < _
       exactRange.Start Then Exit Function
    placeholderRange.Start = _
        placeholderRange.End - Len(expectedToken)
    Set VTNativeEquationArrayPlaceholderRange = _
        placeholderRange.Duplicate
End Function

Public Function VTNativeEquationNumberIsInsideMath( _
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
    Set exactRange = VTNativeMathRoot(formulaRange.OMaths(1)).Range.Duplicate
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

Public Function VTNativeEquationArrayNumberRange( _
    ByVal formulaRange As Range, _
    ByVal numberField As Field) As Range

    Dim exactRange As Range
    Dim markerRange As Range
    Dim fieldStart As Long
    Dim fieldEnd As Long

    If formulaRange Is Nothing Or numberField Is Nothing Then Exit Function
    If formulaRange.OMaths.Count <> 1 Then Exit Function
    Set exactRange = VTNativeMathRoot(formulaRange.OMaths(1)).Range.Duplicate
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

Public Function VTNativeEquationNumberBookmarkIsCompatible( _
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
    Set exactRange = VTNativeMathRoot(formulaRange.OMaths(1)).Range.Duplicate
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

Public Function VTNativeEquationSequenceHelperField( _
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

Public Function VTInsertEquationHelperParagraphAfter( _
    ByVal formulaParagraph As Range) As Range

    Dim documentObject As Document
    Dim insertionRange As Range
    Dim helperParagraph As Range
    Dim helperStart As Long
    Dim insertErrorNumber As Long
    Dim insertErrorDescription As String

    If formulaParagraph Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The Equation helper insertion paragraph is missing."
    End If
    Set documentObject = formulaParagraph.Document
    helperStart = formulaParagraph.End

    ' Keep the existing fast/normal Word operation first. Some real SVG-backed
    ' image paragraphs on Word for Mac become temporarily non-editable after the
    ' displaced-helper repair and raise runtime error 5904 here, even though a
    ' collapsed insertion point at the paragraph mark is writable. Only that
    ' exceptional case uses the fallback; normal sequential insertion keeps the
    ' same one-call path and its existing performance characteristics.
    On Error GoTo InsertParagraphFailed
    Set insertionRange = formulaParagraph.Duplicate
    insertionRange.InsertParagraphAfter
    On Error GoTo 0
    GoTo ResolveHelper

InsertParagraphFailed:
    insertErrorNumber = Err.Number
    insertErrorDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If insertErrorNumber <> 5904 Then
        Err.Raise insertErrorNumber, "VisualTeX Equation helper", _
            insertErrorDescription
    End If
    If helperStart <= formulaParagraph.Start Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The Equation helper paragraph has no writable boundary."
    End If
    Set insertionRange = documentObject.Range( _
        Start:=helperStart - 1, End:=helperStart - 1)
    insertionRange.Text = vbCr

ResolveHelper:
    Set helperParagraph = documentObject.Range( _
        Start:=helperStart, End:=helperStart).Paragraphs(1).Range.Duplicate
    Set VTInsertEquationHelperParagraphAfter = helperParagraph
End Function

Private Sub VTConfigureNativeEquationArrayParagraph( _
    ByVal paragraphRange As Range)

    If paragraphRange Is Nothing Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The native Equation array paragraph is missing."
    End If
    ' Numbering changes layout, not the user's paragraph style. Caption can
    ' carry a theme accent colour, bold weight and a different font size.
    ' Applying it here recolours an otherwise black native equation.
    With paragraphRange.ParagraphFormat
        If .Alignment <> wdAlignParagraphCenter Then _
            .Alignment = wdAlignParagraphCenter
        If .LeftIndent <> 0! Then .LeftIndent = 0!
        If .RightIndent <> 0! Then .RightIndent = 0!
        If .FirstLineIndent <> 0! Then .FirstLineIndent = 0!
        If .SpaceBefore <> 0! Then .SpaceBefore = 0!
        If .SpaceAfter <> 0! Then .SpaceAfter = 0!
        If .LineSpacingRule <> wdLineSpaceSingle Then _
            .LineSpacingRule = wdLineSpaceSingle
        If .KeepWithNext <> False Then .KeepWithNext = False
        If .KeepTogether <> False Then .KeepTogether = False
        If .PageBreakBefore <> False Then .PageBreakBefore = False
        If .WidowControl <> True Then .WidowControl = True
        If .TabStops.Count > 0 Then .TabStops.ClearAll
    End With
End Sub

Public Sub VTRefreshParagraphEquationBookmarks( _
    ByVal documentObject As Document, _
    ByVal sequenceField As Field, _
    ByVal formulaId As String, _
    Optional ByVal preserveFormattedSequence As Boolean = False)

    Dim formulaParagraph As Range
    Dim sequenceParagraph As Range
    Dim numberRange As Range
    Dim formulaRange As Range
    Dim captionRange As Range
    Dim fieldRange As Range
    Dim visibleNumberField As Field
    Dim candidateField As Field
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim captionBookmarkName As String
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim numberSize As Single
    Dim numberFontName As String
    Dim numberPosition As Long
    Dim numberStart As Long
    Dim requiresImageNumberRebuild As Boolean
    Dim expectedNumber As String
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
    If preserveFormattedSequence And _
       documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        Set sequenceParagraph = _
            sequenceField.Result.Paragraphs(1).Range.Duplicate
    Else
        Set sequenceParagraph = _
            VTRefreshFormattedSequenceBookmark( _
                documentObject, sequenceField, formulaId)
        Set sequenceField = VTEquationSequenceFieldForBookmark( _
            documentObject, sequenceBookmarkName)
    End If
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The formatted Equation number lost its native SEQ field."
    End If
    Set sequenceParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate

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
           sequenceParagraph.Start <> formulaParagraph.End Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The image Equation SEQ helper is not immediately after its visible formula row."
        End If
        VTConfigureNumberedEquationParagraph formulaParagraph
        ' The caller (or the prologue above) already formatted this same SEQ.
        ' Do not delete/re-add VT_N_ and reformat its frame a second time.
        Set sequenceField = VTEquationSequenceFieldForBookmark( _
            documentObject, sequenceBookmarkName)
        Set formulaRange = VTNumberedFormulaRangeForId( _
            documentObject, formulaId)
        If sequenceField Is Nothing Or formulaRange Is Nothing Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The image Equation external SEQ identities could not be refreshed."
        End If
        Set visibleNumberField = VTImageEquationReferenceField( _
            formulaRange, formulaId)
        If Not visibleNumberField Is Nothing Then
            visibleNumberField.Update
            Set numberRange = VTImageEquationNumberRange( _
                formulaRange, visibleNumberField)
        End If
        If visibleNumberField Is Nothing Or numberRange Is Nothing Or _
           VTComparableEquationNumberText( _
               VTEquationSequenceResultText(visibleNumberField)) <> _
           VTComparableEquationNumberText( _
               VTEquationNumberTextForFormula(documentObject, formulaId)) Then
            Set numberRange = VTWriteImageEquationReferenceNumber( _
                documentObject, formulaRange, formulaId)
            Set formulaRange = VTNumberedFormulaRangeForId( _
                documentObject, formulaId)
            Set visibleNumberField = VTImageEquationReferenceField( _
                formulaRange, formulaId)
            If visibleNumberField Is Nothing Then
                Err.Raise vbObjectError + 7560, "VisualTeX", _
                    "Word did not rebuild the visible image Equation REF."
            End If
        End If
        VTImageFormulaNumberFormat formulaRange.InlineShapes(1), _
            numberFontName, numberSize
        VTFormatVisibleEquationReference _
            documentObject, visibleNumberField, numberRange, _
            numberFontName, numberSize
        numberPosition = 0
        VTFormatHiddenEquationParagraph sequenceParagraph
        Set sequenceParagraph = _
            sequenceField.Result.Paragraphs(1).Range.Duplicate
        Set captionRange = sequenceParagraph.Paragraphs(1).Range.Duplicate
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
        If VTComparableEquationNumberText( _
               VTEquationSequenceResultText(visibleNumberField)) <> _
           VTComparableEquationNumberText( _
               VTEquationNumberTextForFormula(documentObject, formulaId)) Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The visible native Equation REF does not match its external SEQ."
        End If
        Set numberRange = VTNativeEquationArrayNumberRange( _
            formulaRange, visibleNumberField)
        If numberRange Is Nothing Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "Word did not expose the native Equation array number range."
        End If
        numberSize = VTVisibleEquationNumberFontSize(documentObject)
        numberFontName = VT_WORD_EQUATION_NUMBER_FONT_NAME
        numberPosition = 0
        VTFormatHiddenEquationParagraph sequenceParagraph
        Set sequenceParagraph = _
            sequenceField.Result.Paragraphs(1).Range.Duplicate
        Set captionRange = sequenceParagraph.Paragraphs(1).Range.Duplicate
    Else
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The Equation does not contain exactly one supported visible formula."
    End If

    If numberSize <= 0! Or _
       numberSize > CSng(VT_WORD_MAX_FORMULA_FONT_SIZE_PT) Then
        numberSize = numberRange.Font.Size
    End If
    If numberSize <= 0! Or _
       numberSize > CSng(VT_WORD_MAX_FORMULA_FONT_SIZE_PT) Then
        numberSize = VTVisibleEquationNumberFontSize(documentObject)
    End If
    If numberPosition = wdUndefined Or numberPosition < -512 Or _
       numberPosition > 512 Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The Equation number has an invalid vertical position."
    End If

    VTSetEquationNumberBookmarkExact _
        documentObject, formulaId, numberRange
    With numberRange.Font
        .Hidden = False
        .Position = numberPosition
        If Not VT_WORD_KEEP_NUMBER_FORMAT Then
            .Color = wdColorAutomatic
            .Name = numberFontName
            .NameAscii = numberFontName
            .NameOther = numberFontName
            .Size = numberSize
        End If
    End With
    If formulaRange.InlineShapes.Count = 1 Then
        numberPosition = VTCalculateStaticImageEquationNumberPosition( _
            formulaRange.InlineShapes(1), numberRange)
    End If

    If documentObject.Bookmarks.Exists(captionBookmarkName) Then
        documentObject.Bookmarks(captionBookmarkName).Delete
    End If
    documentObject.Bookmarks.Add _
        name:=captionBookmarkName, Range:=captionRange
End Sub

Public Sub VTVerifyParagraphEquationNumberIntegrity( _
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
    Dim expectedOrdinalText As String
    Dim expectedNumberText As String
    Dim diagnosticText As String
    Dim helperOwnsSequence As Boolean
    Dim helperAfterFormula As Boolean
    Dim numberBoundaryExact As Boolean
    Dim numberOutsideTable As Boolean
    Dim numberInFormulaParagraph As Boolean
    Dim numberTextMatches As Boolean
    Dim macroButtonValid As Boolean
    Dim captionBookmarkValid As Boolean
    Dim restartLevel As Long

    Set sequenceField = VTPureNativeNumberField(formulaRange)
    Set visibleNumberField = VTImagePlaceRef(formulaRange)
    If Not sequenceField Is Nothing Or Not visibleNumberField Is Nothing Then
        If formulaRange.Paragraphs.Count <> 1 Or _
           formulaRange.Paragraphs(1).Range.Frames.Count <> 0 Then
            Err.Raise vbObjectError + 7560, "VisualTeX", "The equation number left its ordinary paragraph."
        End If
        If Not visibleNumberField Is Nothing Then
            Set numberRange = VTCurrentImageNumberRange(visibleNumberField)
            If numberRange Is Nothing Then Err.Raise vbObjectError + 7560
            If numberRange.Start < formulaRange.End Then Err.Raise vbObjectError + 7560
        Else
            Set numberRange = sequenceField.Result.Duplicate
        End If
        If Len(Trim$(numberRange.Text)) = 0 Then
            Err.Raise vbObjectError + 7560, "VisualTeX", "The native sequence result is empty."
        End If
        Exit Sub
    End If
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
       formulaRange.Start < formulaParagraph.Start Or _
       formulaRange.End > formulaParagraph.End Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The numbered formula is not inside one ordinary Word paragraph" & _
            " [formula=" & CStr(formulaRange.Start) & "-" & _
                CStr(formulaRange.End) & _
            "; paragraph=" & CStr(formulaParagraph.Start) & "-" & _
                CStr(formulaParagraph.End) & _
            "; table=" & CStr(formulaParagraph.Information( _
                wdWithInTable)) & "]."
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
    restartLevel = VTEquationNumberingRestartLevel(documentObject)
    expectedOrdinalText = CStr(expectedOrdinal)
    If Not VTEquationSequenceFieldHasOrdinal( _
       sequenceField, VTNativeEquationLabelName(), _
       expectedOrdinal, restartLevel) Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The Equation SEQ result is incorrect" & _
            " [code=" & sequenceField.Code.Text & _
            "; result=" & VTEquationSequenceResultText(sequenceField) & _
            "; ordinal=" & expectedOrdinalText & _
            "; restartLevel=" & CStr(restartLevel) & "]."
    End If
    expectedNumberText = VTEquationNumberTextForFormula( _
        documentObject, formulaId)
    If Len(expectedNumberText) = 0 Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The Equation has no formatted visible number."
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
        If visibleNumberField Is Nothing Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The image Equation visible number REF is missing."
        End If
        Set expectedNumberRange = VTImageEquationNumberRange( _
            formulaRange, visibleNumberField)
        If expectedNumberRange Is Nothing Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The image Equation visible number range is missing."
        End If
        helperOwnsSequence = _
            VTHelperParagraphOwnsNativeEquationSequence(helperParagraph)
        helperAfterFormula = _
            helperParagraph.Start = formulaParagraph.End
        numberBoundaryExact = _
            numberRange.Start = expectedNumberRange.Start And _
            numberRange.End = expectedNumberRange.End
        numberOutsideTable = _
            Not numberRange.Information(wdWithInTable)
        numberInFormulaParagraph = _
            numberRange.Paragraphs(1).Range.Start = formulaParagraph.Start
        numberTextMatches = _
            numberRange.Text = "(" & expectedNumberText & ")" And _
            VTComparableEquationNumberText( _
                VTEquationSequenceResultText(visibleNumberField)) = _
            VTComparableEquationNumberText(expectedNumberText)
        macroButtonValid = _
            VTParagraphHasSingleVisualTeXImageMacroButton(formulaParagraph)
        captionBookmarkValid = _
            VTEquationCaptionBookmarkIsCollapsedInParagraph( _
                captionRange, helperParagraph)
        If Not helperOwnsSequence Or Not helperAfterFormula Or _
           Not numberBoundaryExact Or Not numberOutsideTable Or _
           Not numberInFormulaParagraph Or Not numberTextMatches Or _
           Not macroButtonValid Or Not captionBookmarkValid Then
            diagnosticText = _
                "helperOwns=" & CStr(helperOwnsSequence) & _
                "; helperAdjacent=" & CStr(helperAfterFormula) & _
                "; numberExact=" & CStr(numberBoundaryExact)
            diagnosticText = diagnosticText & _
                "; outsideTable=" & CStr(numberOutsideTable) & _
                "; sameParagraph=" & CStr(numberInFormulaParagraph) & _
                "; textMatches=" & CStr(numberTextMatches)
            diagnosticText = diagnosticText & _
                "; macroButton=" & CStr(macroButtonValid) & _
                "; caption=" & CStr(captionBookmarkValid) & _
                "; numberText=[" & numberRange.Text & "]" & _
                "; expected=[(" & expectedNumberText & ")]"
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The image Equation external SEQ/visible REF scaffold is incomplete" & _
                " [" & diagnosticText & "]."
        End If
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
       VTComparableEquationNumberText( _
           VTEquationSequenceResultText(visibleNumberField)) <> _
       VTComparableEquationNumberText(expectedNumberText) Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The native Equation visible REF is invalid" & _
            " [actual=" & _
                VTEquationSequenceResultText(visibleNumberField) & _
            "; expected=" & expectedNumberText & _
            "; insideMath=" & CStr( _
                VTNativeEquationNumberIsInsideMath( _
                    formulaRange, visibleNumberField)) & "]."
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

Public Sub VTFinalizeParagraphEquationNumber( _
    ByVal documentObject As Document, _
    ByVal formulaRange As Range, _
    ByVal formulaId As String, _
    Optional ByVal deferReconcile As Boolean = False, _
    Optional ByVal sequenceOrderChanged As Boolean = True)

    Dim formulaParagraph As Range
    Dim sequenceField As Field
    Dim equationLabelName As String
    Dim sequenceBookmarkName As String
    Dim sequenceOrdinal As Long
    Dim restartLevel As Long
    Dim numberingMode As String
    Dim fieldAnchor As Long
    Dim preserveFormattedSequence As Boolean

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

    If Not VTPureNativeNumberField(formulaRange) Is Nothing Then
        VTUpdateCurrentWordNumberFields documentObject
        Exit Sub
    End If
    If Not VTImagePlaceRef(formulaRange) Is Nothing Then
        VTUpdateCurrentWordNumberFields documentObject
        Exit Sub
    End If
    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    If formulaRange.InlineShapes.Count = 1 And _
       formulaRange.OMaths.Count = 0 Then
        Set sequenceField = VTNativeEquationSequenceHelperField( _
            documentObject, formulaId)
    ElseIf formulaRange.InlineShapes.Count = 0 And _
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

    ' A freshly inserted native SEQ already evaluates itself against preceding
    ' Equation fields, and an edited formula keeps its existing result. Read that
    ' local result first instead of counting every Field from the document start.
    ' The document-wide ordinal scan remains only as a compatibility fallback.
    sequenceOrdinal = VTFirstPositiveIntegerInText( _
        VTEquationSequenceResultText(sequenceField))
    If sequenceOrdinal < 1 Then
        sequenceOrdinal = VTEquationSequenceOrdinal( _
            documentObject, sequenceField, equationLabelName)
    End If
    If sequenceOrdinal < 1 Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The Equation has no usable sequence result."
    End If
    numberingMode = VTEquationNumberingMode(documentObject)
    restartLevel = VTEquationNumberingRestartLevel(documentObject)
    If Not VTEquationSequenceFieldHasOrdinal( _
       sequenceField, equationLabelName, sequenceOrdinal, restartLevel) Then
        VTApplyEquationSequenceOrdinal _
            sequenceField, equationLabelName, sequenceOrdinal, restartLevel
    End If
    Set sequenceField = VTResolveEquationSequenceFieldNear( _
        documentObject, fieldAnchor, 64)
    preserveFormattedSequence = Not sequenceOrderChanged
    If Not preserveFormattedSequence Then
        preserveFormattedSequence = _
            numberingMode = VT_WORD_NUMBERING_MODE_SEQUENCE
    End If
    VTRefreshParagraphEquationBookmarks _
        documentObject, sequenceField, formulaId, _
        preserveFormattedSequence
    Set sequenceField = VTEquationSequenceFieldForBookmark( _
        documentObject, sequenceBookmarkName)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "Word lost the local Equation SEQ after refreshing its number."
    End If
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

    ' Editing an existing numbered formula does not change sequence order and can
    ' keep the local fast path. The first and only managed Equation is also fully
    ' resolved by the local refresh above: its heading prefix was calculated from
    ' the current OutlineLevel scope, while there are no other managed SEQ/REF
    ' identities to replay. Later chapter/section appends retain the shared scan
    ' so custom outline boundaries and following formulas remain exact.
    If Not sequenceOrderChanged Or _
       Not VTHasOtherManagedEquationNumber( _
           documentObject, formulaId) Or _
       (numberingMode = VT_WORD_NUMBERING_MODE_SEQUENCE And _
        VTCanUseEquationTailFastPath( _
            documentObject, sequenceField)) Then
        Set formulaRange = VTNumberedFormulaRangeForId( _
            documentObject, formulaId)
        If formulaRange Is Nothing Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "Word lost the locally finalized Equation."
        End If
        VTVerifyParagraphEquationNumberIntegrity _
            formulaRange, formulaId, sequenceOrdinal
        VTEnsureOrphanWatchScheduled
        Exit Sub
    End If

    ' Word for Mac can leave a freshly inserted middle SEQ at result 1 even when
    ' its physical ordinal is much later in the document. Updating only the suffix
    ' does not reliably rebuild Word's internal flowing-SEQUENCE state. Middle
    ' insertion is intentionally the conservative path: replay the managed SEQ
    ' sequence from the beginning and force every flowing SEQ to evaluate in that
    ' order, so the new field receives its positional ordinal before every following
    ' formula and body REF is refreshed. True tail appends never reach this branch
    ' and retain the O(1) local fast path above.
    VTReconcileEquationNumbers documentObject, -1, True
    Set sequenceField = VTEquationSequenceFieldForBookmark( _
        documentObject, sequenceBookmarkName)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "Word lost the Equation SEQ during reconciliation."
    End If
    sequenceOrdinal = VTFirstPositiveIntegerInText( _
        VTEquationSequenceResultText(sequenceField))
    If sequenceOrdinal < 1 Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The reconciled Equation SEQ has no local ordinal."
    End If
    VTRefreshParagraphEquationBookmarks _
        documentObject, sequenceField, formulaId, True
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

Public Function VTFindEquationSequenceField( _
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

Public Function VTEquationLayoutWidth( _
    ByVal paragraphRange As Range) As Single

    Dim sectionObject As Section
    Dim textWidth As Single
    Dim columnCount As Long

    If paragraphRange Is Nothing Then
        Err.Raise vbObjectError + 7425, "VisualTeX", _
            "The Equation layout paragraph is missing."
    End If
    If VT_WORD_DISPLAY_NUMBERING_LAYOUT_WIDTH > 0! Then
        VTEquationLayoutWidth = VT_WORD_DISPLAY_NUMBERING_LAYOUT_WIDTH
        Exit Function
    End If
    On Error Resume Next
    Set sectionObject = paragraphRange.Sections(1)
    columnCount = sectionObject.PageSetup.TextColumns.Count
    textWidth = sectionObject.PageSetup.TextColumns.Width
    ' After replacing OMath, Mac Word can temporarily return wdUndefined
    ' for the collection width even in a single-column section. Never use
    ' that sentinel as a layout dimension. Resolve the one actual column
    ' first; page margins are a valid fallback only for a single column.
    If (textWidth <= 0! Or textWidth = wdUndefined) And columnCount = 1 Then
        textWidth = sectionObject.PageSetup.TextColumns(1).Width
    End If
    If (textWidth <= 0! Or textWidth = wdUndefined) And columnCount = 1 Then
        textWidth = sectionObject.PageSetup.PageWidth - _
            sectionObject.PageSetup.LeftMargin - _
            sectionObject.PageSetup.RightMargin
    End If
    On Error GoTo 0
    If textWidth <= 2! Or textWidth = wdUndefined Then
        Err.Raise vbObjectError + 7425, "VisualTeX", _
            "Word returned an invalid text width for equation numbering."
    End If
    VTEquationLayoutWidth = textWidth
End Function

' Word for Mac can refuse a paragraph's Style once right after the picture in
' it was resized ("Method 'Style' of object 'Range' failed") and accept it on
' the next try.
Private Sub VTApplyParagraphStyle( _
    ByVal paragraphRange As Range, _
    ByVal paragraphStyle As Style)

    Dim attempt As Long
    On Error Resume Next
    For attempt = 1 To 3
        Err.Clear
        paragraphRange.Style = paragraphStyle
        If Err.Number = 0 Then Exit Sub
    Next attempt
    On Error GoTo 0
    paragraphRange.Style = paragraphStyle
End Sub

Public Sub VTConfigureNumberedEquationParagraph( _
    ByVal paragraphRange As Range)

    Dim textWidth As Single
    Dim imageEquationStyle As Style
    Dim typographyAlignment As Long

    typographyAlignment = paragraphRange.ParagraphFormat.BaseLineAlignment
    textWidth = VTEquationLayoutWidth(paragraphRange)
    If paragraphRange.InlineShapes.Count = 1 And _
       paragraphRange.OMaths.Count = 0 Then
        ' Give Word a real following-paragraph contract instead of repairing
        ' Return after the fact from SelectionChange. Direct center/right tabs
        ' remain paragraph-local so different sections can keep different widths,
        ' while NextParagraphStyle makes Return create a clean Normal paragraph.
        Set imageEquationStyle = _
            VTEnsureNumberedImageParagraphStyle(paragraphRange.Document)
        VTApplyParagraphStyle paragraphRange, imageEquationStyle
    ElseIf paragraphRange.Style = wdStyleCaption Then
        ' Preserve the pre-existing native compatibility behavior.
        paragraphRange.Style = wdStyleNormal
    End If
    With paragraphRange.ParagraphFormat
        .BaseLineAlignment = typographyAlignment
        .Alignment = wdAlignParagraphJustify
        .LeftIndent = 0!
        .RightIndent = 0!
        .FirstLineIndent = 0!
        .SpaceBefore = 0!
        .SpaceAfter = 0!
        .LineSpacingRule = wdLineSpaceSingle
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .WidowControl = False
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

Public Function VTImportNativeNumberShell( _
    ByVal sourceRange As Range, ByVal formulaId As String) As Range

    Dim shellDocument As Document
    Dim parsedDocument As Document
    Dim targetDocument As Document
    Dim shellRange As Range
    Dim replacementRange As Range
    Dim shellMath As OMath
    Dim sourceXml As String
    Dim shellXml As String
    Dim sourceBody As String
    Dim shellBody As String
    Dim sourceBodyStart As Long
    Dim sourceBodyEnd As Long
    Dim shellBodyStart As Long
    Dim shellBodyEnd As Long
    Dim tokenStart As Long
    Dim runStart As Long
    Dim runEnd As Long
    Dim sourceStart As Long
    Dim parsedLength As Long
    Dim stagingPath As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed
    Set targetDocument = sourceRange.Document
    sourceStart = sourceRange.Start
    sourceXml = sourceRange.WordOpenXML
    sourceBody = VTNativeMathXmlBody(sourceXml, sourceBodyStart, sourceBodyEnd)

    ' BuildUp sees only this disposable linear placeholder, never the user's
    ' n-ary/delimiter/noBar tree, and never a live REF or SEQ field.
    Set shellDocument = Documents.Add(Visible:=False)
    Set shellRange = shellDocument.Content.Duplicate
    shellRange.Text = "VTNUMBERBODY#(0)"
    shellRange.End = shellDocument.Content.End - 1
    Set shellRange = shellDocument.OMaths.Add(shellRange)
    Set shellMath = shellRange.OMaths(1)
    shellMath.BuildUp
    shellMath.Type = wdOMathDisplay
    shellMath.Justification = wdOMathJcCenterGroup
    shellXml = shellMath.Range.WordOpenXML
    shellBody = VTNativeMathXmlBody(shellXml, shellBodyStart, shellBodyEnd)
    tokenStart = InStr(1, shellBody, "<m:t>VTNUMBERBODY</m:t>", vbBinaryCompare)
    If tokenStart = 0 Then GoTo InvalidShell
    runStart = InStrRev(shellBody, "<m:r>", tokenStart, vbBinaryCompare)
    runEnd = InStr(tokenStart, shellBody, "</m:r>", vbBinaryCompare)
    If runStart = 0 Or runEnd = 0 Then GoTo InvalidShell
    ' Keep the source Flat OPC package and its namespaces. Replace only the
    ' OMath body with Word's built shell containing the exact source XML tree.
    shellBody = Left$(shellBody, runStart - 1) & sourceBody & _
        Mid$(shellBody, runEnd + Len("</m:r>"))
    sourceXml = Left$(sourceXml, sourceBodyStart - 1) & shellBody & _
        Mid$(sourceXml, sourceBodyEnd)
    stagingPath = VTApplicationSupportRoot() & _
        "/number-shell-" & VTNewUuidV4() & ".xml"
    VTWriteTextAtomic stagingPath, sourceXml
    Set parsedDocument = Documents.Open( _
        FileName:=stagingPath, ConfirmConversions:=False, ReadOnly:=True, _
        AddToRecentFiles:=False, Visible:=False)
    If parsedDocument.OMaths.Count <> 1 Or _
       parsedDocument.InlineShapes.Count <> 0 Or _
       parsedDocument.Tables.Count <> 0 Then GoTo InvalidShell
    Set shellRange = VTNativeMathRoot(parsedDocument.OMaths(1)).Range.Duplicate
    If VTNativeEquationArrayMarkerRange(shellRange) Is Nothing Then GoTo InvalidShell
    parsedLength = shellRange.End - shellRange.Start
    targetDocument.Activate
    ' Transfer only the parsed OMath, never its paragraph mark or helper. The
    ' caller's conversion/commit transaction owns rollback of this replacement.
    Set replacementRange = sourceRange.Duplicate
    replacementRange.FormattedText = shellRange.FormattedText
    Set replacementRange = targetDocument.Range( _
        Start:=sourceStart, End:=sourceStart + parsedLength)
    If replacementRange.OMaths.Count <> 1 Then GoTo InvalidShell
    Set VTImportNativeNumberShell = VTNativeMathRoot(replacementRange.OMaths(1)).Range.Duplicate
    parsedDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set parsedDocument = Nothing
    shellDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set shellDocument = Nothing
    On Error Resume Next
    Kill stagingPath
    On Error GoTo 0
    Exit Function
InvalidShell:
    Err.Raise vbObjectError + 7563, "VisualTeX", _
        "Word did not preserve the isolated professional Equation number shell."
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Not parsedDocument Is Nothing Then parsedDocument.Close SaveChanges:=wdDoNotSaveChanges
    If Not shellDocument Is Nothing Then shellDocument.Close SaveChanges:=wdDoNotSaveChanges
    If Len(stagingPath) > 0 Then Kill stagingPath
    targetDocument.Activate
    On Error GoTo 0
    Err.Raise errorNumber, "VisualTeX number shell", errorDescription
End Function

Public Function VTEnsureNativeEquationNumber( _
    ByVal equationRange As Range, _
    ByVal renderedHeightPoints As Double, _
    ByVal formulaId As String, _
    ByVal captionText As String, _
    ByRef numberCreated As Boolean, _
    Optional ByRef performancePhases As Variant) As Range
    If equationRange Is Nothing Then Err.Raise vbObjectError + 7606
    numberCreated = (VTPureNativeNumberField(equationRange) Is Nothing)
    Set VTEnsureNativeEquationNumber = VTWritePureNativeNumber(equationRange, formulaId)
    If Not IsMissing(performancePhases) Then performancePhases = Array(0#, 0#, 0#, 0#, 0#, 0#)
End Function

Public Sub VTConfigureNumberedDisplayTable(ByVal layoutTable As Table)
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

Public Sub VTEnsureEquationNumberFields( _
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
        Text:=VTVisibleEquationReferenceFieldText( _
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

Public Function VTVisibleEquationNumberFontSize( _
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

Public Sub VTVerifyEquationNumberFieldIntegrity( _
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

Public Function VTBookmarkNameInFieldCode( _
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

Public Sub VTReconcileEquationNumbers( _
    ByVal documentObject As Document, _
    Optional ByVal changedFrom As Long = -1, _
    Optional ByVal forceFlowingSequenceReplay As Boolean = False)
    If documentObject Is Nothing Then Exit Sub
    VTPinEquationNumbering documentObject
    VTUpdateCurrentWordNumberFields documentObject
End Sub

' Native Equation numbers are SEQ (and STYLEREF) fields inside their equation.
' The document's fields are cheap to enumerate on Word for Mac, so the
' equations that can carry a number are found from them instead of from
' Document.OMaths. Returns, once each and in document order, the root Range of
' every native equation holding a SEQ field (or a STYLEREF field when
' includeStyleRef is True) whose field code starts in [searchStart, searchEnd).
Public Function VTNumberFieldNativeMathRoots( _
    ByVal documentObject As Document, _
    ByVal searchStart As Long, _
    ByVal searchEnd As Long, _
    Optional ByVal includeStyleRef As Boolean = False) As Collection

    Dim result As Collection
    Dim scanRange As Range
    Dim candidate As Field
    Dim rootRange As Range
    Dim fieldType As Long
    Dim lastStart As Long

    Set result = New Collection
    Set VTNumberFieldNativeMathRoots = result
    If documentObject Is Nothing Then Exit Function
    If searchStart < documentObject.Content.Start Then _
        searchStart = documentObject.Content.Start
    If searchEnd > documentObject.Content.End Then _
        searchEnd = documentObject.Content.End
    If searchEnd <= searchStart Then Exit Function
    Set scanRange = documentObject.Range(Start:=searchStart, End:=searchEnd)
    lastStart = -1
    For Each candidate In scanRange.Fields
        fieldType = candidate.Type
        If fieldType = wdFieldSequence Or _
           (includeStyleRef And fieldType = wdFieldStyleRef) Then
            If candidate.Code.Start >= searchStart And _
               candidate.Code.Start < searchEnd Then
                Set rootRange = VTFieldNativeMathRoot(candidate)
                If Not rootRange Is Nothing Then
                    If rootRange.Start > lastStart Then
                        result.Add rootRange
                        lastStart = rootRange.Start
                    End If
                End If
            End If
        End If
    Next candidate
End Function

' The Word font of an image formula's Equation number: the formula's own
' letter font. KaTeX / Computer Modern has no Word font, and a formula whose
' letter font is not known yet keeps Cambria Math.
Public Function VTEquationNumberFontNameForFormula( _
    ByVal documentObject As Document, _
    ByVal formulaId As String) As String

    Dim letterFont As String
    VTEquationNumberFontNameForFormula = VT_WORD_EQUATION_NUMBER_FONT_NAME
    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then Exit Function
    If Not VTTryGetDocumentVariable(documentObject, _
        VTWordImageLetterFontVariableName(formulaId), letterFont) Then Exit Function
    Select Case letterFont
        Case "times": VTEquationNumberFontNameForFormula = "Times New Roman"
        Case "stix": VTEquationNumberFontNameForFormula = "STIX Two Text"
        Case "palatino": VTEquationNumberFontNameForFormula = "Palatino"
        Case "helvetica": VTEquationNumberFontNameForFormula = "Helvetica Neue"
    End Select
End Function

' Painted top and bottom of "(", ")" and the digits, and the line ascent and
' descent, in em, for each Equation number font (CoreText; the Cambria Math
' values are those of Word's own copy of the font).
Public Sub VTEquationNumberFontMetrics( _
    ByVal fontName As String, _
    ByRef inkTopEm As Double, _
    ByRef inkBottomEm As Double, _
    ByRef fontTopEm As Double, _
    ByRef fontDescentEm As Double)

    Select Case LCase$(Trim$(fontName))
        Case "times new roman"
            inkTopEm = 0.69434: inkBottomEm = -0.21387
            fontTopEm = 0.89111: fontDescentEm = 0.21631
        Case "stix two text"
            inkTopEm = 0.73: inkBottomEm = -0.186
            fontTopEm = 0.76199: fontDescentEm = 0.23801
        Case "palatino"
            inkTopEm = 0.77051: inkBottomEm = -0.16602
            fontTopEm = 0.82275: fontDescentEm = 0.27734
        Case "helvetica neue"
            inkTopEm = 0.731: inkBottomEm = -0.197
            fontTopEm = 0.952: fontDescentEm = 0.213
        Case Else
            inkTopEm = 0.70556640625#: inkBottomEm = -0.220703125#
            fontTopEm = 0.9501953125#: fontDescentEm = 0.22216796875#
    End Select
End Sub

Private Function VTNativeSequenceInstruction( _
    ByVal documentObject As Document, _
    Optional ByVal hiddenIncrement As Boolean = False, _
    Optional ByVal currentValue As Boolean = False, _
    Optional ByVal restartAtOne As Boolean = False) As String
    Dim instruction As String
    instruction = "SEQ " & VTEquationSequenceFieldText(VTNativeEquationLabelName())
    If currentValue Then
        instruction = instruction & " \c"
    ElseIf restartAtOne Then
        instruction = instruction & " \r 1"
    End If
    If hiddenIncrement Then
        instruction = instruction & " \h"
    Else
        instruction = instruction & " \* ARABIC"
    End If
    VTNativeSequenceInstruction = instruction
End Function

Public Function VTInsertFieldBefore( _
    ByVal beforeRange As Range, ByVal instruction As String, _
    Optional ByVal preserveFormatting As Boolean = False) As Field
    Dim insertion As Range
    Set insertion = beforeRange.Duplicate
    insertion.Collapse wdCollapseStart
    Set VTInsertFieldBefore = insertion.Fields.Add( _
        Range:=insertion, Type:=wdFieldEmpty, Text:=instruction, _
        PreserveFormatting:=preserveFormatting)
End Function

Public Function VTPureNativeNumberField(ByVal formulaRange As Range) As Field
    Dim candidate As Field
    Dim exact As Range
    Dim marker As Range
    Dim xml As String
    Dim mathFunction As OMathFunction
    Dim hasArray As Boolean
    If formulaRange Is Nothing Then Exit Function
    If formulaRange.OMaths.Count <> 1 Then Exit Function
    Set exact = VTNativeMathRoot(formulaRange.OMaths(1)).Range.Duplicate
    If exact.Fields.Count = 0 Then Exit Function
    ' Word may split the # marker across math runs after an edit. Require the
    ' native equation-array shell, but do not depend on one exact XML spelling.
    ' The native object model identifies the canonical array without exporting
    ' a full Flat OPC package for every equation in a dense chapter scan.
    For Each mathFunction In exact.OMaths(1).Functions
        If mathFunction.Type = wdOMathFunctionEqArray Then
            hasArray = True
            Exit For
        End If
    Next mathFunction
    If Not hasArray Then
        ' Retain the former XML check for host builds with incomplete Functions.
        xml = exact.WordOpenXML
        If InStr(1, xml, "<m:eqArr", vbBinaryCompare) = 0 Or _
           InStr(1, xml, "#", vbBinaryCompare) = 0 Then Exit Function
    End If
    Set marker = VTNativeEquationArrayMarkerRange(exact)
    If marker Is Nothing Then Exit Function
    For Each candidate In exact.Fields
        If candidate.Type = wdFieldSequence And _
           VTEquationFieldStart(candidate) >= marker.End And _
           VTEquationFieldEnd(candidate) <= exact.End Then
            ' The native Equation caption identifier is localized inside the
            ' saved field code. Word can expose a different CaptionLabels name
            ' after save/reopen, so structural ownership is authoritative: a
            ' SEQ field in the right-hand # slot of one Equation-array is the
            ' native equation number. Requiring the current UI label here made
            ' an existing black number reopen as an unchecked edit Session.
            Set VTPureNativeNumberField = candidate
            Exit Function
        End If
    Next candidate
End Function

Public Function VTPureNativeVisibleNumberRange( _
    ByVal formulaRange As Range, ByVal numberField As Field) As Range

    Dim prefixField As Field
    Dim source As Range

    If formulaRange Is Nothing Or numberField Is Nothing Then Exit Function
    Set prefixField = VTPureNativePrefixField(formulaRange, numberField)
    If prefixField Is Nothing Then
        Set source = numberField.Result.Duplicate
    Else
        Set source = formulaRange.Document.Range( _
            prefixField.Result.Start, numberField.Result.End)
    End If
    source.TextRetrievalMode.IncludeFieldCodes = False
    Set VTPureNativeVisibleNumberRange = source
End Function

Private Function VTPureNativePrefixField( _
    ByVal formulaRange As Range, ByVal numberField As Field) As Field

    Dim exact As Range
    Dim marker As Range
    Dim candidate As Field

    If formulaRange Is Nothing Or numberField Is Nothing Then Exit Function
    If formulaRange.OMaths.Count <> 1 Then Exit Function
    Set exact = VTNativeMathRoot(formulaRange.OMaths(1)).Range.Duplicate
    Set marker = VTNativeEquationArrayMarkerRange(exact)
    If marker Is Nothing Then Exit Function
    For Each candidate In exact.Fields
        If candidate.Type = wdFieldQuote And _
           VTEquationFieldStart(candidate) >= marker.End And _
           candidate.Result.Start <= numberField.Result.Start Then
            Set VTPureNativePrefixField = candidate
            Exit Function
        End If
    Next candidate
End Function

Private Function VTNativePrefixInstruction(ByVal prefixText As String) As String
    VTNativePrefixInstruction = "QUOTE " & Chr$(34) & _
        Replace$(prefixText, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Public Function VTWritePureNativeNumber( _
    ByVal equationRange As Range, _
    ByVal formulaId As String, _
    Optional ByVal headingPrefixOverride As String = "", _
    Optional ByVal headingPrefixResolved As Boolean = False, _
    Optional ByVal restartAtOneOverride As Boolean = False, _
    Optional ByVal restartResolved As Boolean = False) As Range
    Dim doc As Document
    Dim exact As Range
    Dim marker As Range
    Dim placeholder As Range
    Dim field As Field
    Dim prefixField As Field
    Dim source As Range
    Dim level As Long
    Dim headingPrefix As String
    Dim prefixText As String
    Dim restartAtOne As Boolean
    Dim operationStage As String
    On Error GoTo Failed
    operationStage = "resolve"
    Set doc = equationRange.Document
    Set exact = VTNativeMathRoot(equationRange.OMaths(1)).Range.Duplicate
    Set field = VTPureNativeNumberField(exact)
    If field Is Nothing Then
        Set marker = VTNativeEquationArrayMarkerRange(exact)
        If Not marker Is Nothing Then
            ' A freshly imported numbered cache deliberately contains the exact
            ' VisualTeX-owned sentinel shell. Resolve that trusted placeholder before
            ' looking for the legacy REF form. Any other marked equation remains
            ' user-owned and must not have its number overwritten.
            Set placeholder = VTNativeEquationArrayPlaceholderRange(exact)
            If placeholder Is Nothing Then
                For Each field In exact.Fields
                    If field.Type = wdFieldRef Then
                        Set placeholder = doc.Range( _
                            VTEquationFieldStart(field), _
                            VTEquationFieldEnd(field))
                        Exit For
                    End If
                Next field
            End If
            If placeholder Is Nothing Then Err.Raise vbObjectError + 7606, _
                "VisualTeX", "Preserve a manually numbered native equation instead of overwriting its number."
        Else
            operationStage = "import-shell"
            Set exact = VTImportNativeNumberShell(exact, formulaId)
            Set marker = VTNativeEquationArrayMarkerRange(exact)
            If marker Is Nothing Then Err.Raise vbObjectError + 7606, , "Missing native # marker."
            Set placeholder = VTNativeEquationArrayPlaceholderRange(exact, "0")
            If placeholder Is Nothing Then Err.Raise vbObjectError + 7606, , "Missing number placeholder."
        End If
        If placeholder Is Nothing Then Err.Raise vbObjectError + 7606
        operationStage = "insert-seq"
        level = VTNativeHeadingLevelAtRange(exact)
        If level > 0 Then
            If headingPrefixResolved Then
                headingPrefix = headingPrefixOverride
            Else
                headingPrefix = VTEquationHeadingPrefix( _
                    exact, VTEquationNumberingMode(doc))
            End If
            prefixText = headingPrefix & _
                VTEquationNumberingDisplaySeparator(doc)
            If restartResolved Then
                restartAtOne = restartAtOneOverride
            Else
                restartAtOne = VTEquationStartsNumberScope(exact, False)
            End If
        End If
        Set field = doc.Fields.Add(Range:=placeholder, Type:=wdFieldEmpty, _
            Text:=VTNativeSequenceInstruction( _
                doc, False, False, restartAtOne), _
            PreserveFormatting:=False)
        If Len(prefixText) > 0 Then
            Set placeholder = doc.Range(VTEquationFieldStart(field), VTEquationFieldStart(field))
            Set prefixField = VTInsertFieldBefore( _
                placeholder, VTNativePrefixInstruction(prefixText))
            prefixField.Update
        End If
    End If
    operationStage = "resolve-seq-math"
    Set exact = VTNativeMathRoot(field.Result.Paragraphs(1).Range.OMaths(1)).Range.Duplicate
    Set source = VTPureNativeVisibleNumberRange(exact, field)
    If source Is Nothing Then Err.Raise vbObjectError + 7606, _
        "VisualTeX", "The native Equation visible-number range is missing."
    VTRetireExternalSequenceHelper doc, formulaId, source
    field.Update
    Set exact = VTNativeMathRoot(field.Result.Paragraphs(1).Range.OMaths(1)).Range.Duplicate
    exact.OMaths(1).Type = wdOMathDisplay
    exact.OMaths(1).Justification = wdOMathJcCenterGroup
    VTConfigureNativeEquationArrayParagraph exact.Paragraphs(1).Range
    ' These identities exist only inside the legacy transaction. The commit
    ' boundary strips them and all payload variables, including on conversion.
    VTSetNativeFormulaBookmark doc, exact, formulaId
    Set source = VTPureNativeVisibleNumberRange(exact, field)
    VTSetEquationNumberBookmarkExact doc, formulaId, source
    Set VTWritePureNativeNumber = exact.Duplicate
    Exit Function
Failed:
    Err.Raise Err.Number, "VisualTeX native number", operationStage & ": " & Err.Description
End Function

Public Sub VTUpdateCurrentWordNumberFields(ByVal doc As Document)
    Dim candidate As Field
    Dim math As OMath
    ' Field objects retain their ranges when earlier field results change length.
    ' Traverse main-story paragraphs in order; do not freeze integer positions.
    Dim paragraph As Paragraph
    For Each paragraph In doc.Paragraphs
        For Each math In paragraph.Range.OMaths
            For Each candidate In math.Range.Fields
                If candidate.Type = wdFieldSequence Or candidate.Type = wdFieldStyleRef Then candidate.Update
            Next candidate
        Next math
        For Each candidate In paragraph.Range.Fields
            If VTIsImagePlaceRef(candidate) Then
                VTUpdateImagePlaceRef candidate
            ElseIf candidate.Type = wdFieldSequence Or candidate.Type = wdFieldStyleRef Then
                candidate.Update
            End If
        Next candidate
    Next paragraph
    For Each candidate In doc.Fields
        If candidate.Type = wdFieldRef Then candidate.Update
    Next candidate
End Sub

Private Function VTCanBatchNativeNumberFields(ByVal scan As Range) As Boolean
    Dim candidate As Field
    Dim instruction As String
    Dim mathRange As Range
    For Each candidate In scan.Fields
        Select Case candidate.Type
            Case wdFieldSequence
                If candidate.Result.Paragraphs(1).Range.OMaths.Count <> 1 Then Exit Function
                Set mathRange = candidate.Result.Paragraphs(1).Range.OMaths(1).Range
                If VTEquationFieldStart(candidate) < mathRange.Start Or _
                   VTEquationFieldEnd(candidate) > mathRange.End Then Exit Function
                If StrComp(VTSequenceIdentifierFromFieldCode(candidate.Code.Text), _
                   VTNativeEquationLabelName(), vbTextCompare) <> 0 Then Exit Function
            Case wdFieldQuote
                ' Only literal chapter prefixes are safe to reevaluate as part
                ' of this batch. A calculated/user-owned QUOTE stays untouched.
                instruction = Trim$(candidate.Code.Text)
                If StrComp(instruction, VTNativePrefixInstruction(candidate.Result.Text), _
                   vbTextCompare) <> 0 Then Exit Function
            Case wdFieldStyleRef
                If candidate.Result.Paragraphs(1).Range.OMaths.Count <> 1 Then Exit Function
                Set mathRange = candidate.Result.Paragraphs(1).Range.OMaths(1).Range
                If VTEquationFieldStart(candidate) < mathRange.Start Or _
                   VTEquationFieldEnd(candidate) > mathRange.End Then Exit Function
            Case wdFieldRef
                ' Body references are refreshed again after all SEQ sources.
            Case Else
                ' MACROBUTTON image numbers and unrelated fields must retain
                ' their separate chain. Use the scoped individual path instead.
                Exit Function
        End Select
    Next candidate
    VTCanBatchNativeNumberFields = True
End Function

Public Sub VTReconcileNumberingAfterEdit( _
    ByVal doc As Document, ByVal changedFrom As Long, _
    ByVal imageFormula As Boolean)

    Dim shape As InlineShape
    Dim math As OMath
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim result As Range
    Dim scan As Range
    Dim paragraph As Paragraph
    Dim candidate As Field
    Dim sequence As Field
    Dim restartAtOne As Boolean
    Dim exact As Range

    ' Only the next numbered formula can acquire or lose a chapter/section
    ' restart when membership changes. Repair its instruction, then refresh
    ' the affected sequence suffix before acknowledging this edit to the UI.
    If VTEquationNumberingRestartLevel(doc) > 0 Then
        If imageFormula Then
            For Each shape In doc.InlineShapes
                If shape.Range.Start > changedFrom Then
                    If VTTryParseFormulaReference(shape.Title, formulaId, displayMode, numbered) Then
                        If numbered And displayMode = "block" Then
                            restartAtOne = VTEquationStartsNumberScope(shape.Range, True)
                            Set sequence = VTImageSequenceField(VTImagePlaceRef(shape.Range), True)
                            If Not VTEquationSequenceFieldIsCompatible(sequence, _
                               VT_WORD_IMAGE_SEQUENCE_NAME, True, False, restartAtOne) Then
                                Set result = VTWriteSingleParagraphImageNumber(shape, formulaId)
                            End If
                            Exit For
                        End If
                    End If
                End If
            Next shape
        Else
            For Each exact In VTNumberFieldNativeMathRoots( _
                doc, changedFrom, doc.Content.End)
                If exact.Start > changedFrom Then
                    If Not VTPureNativeNumberField(exact) Is Nothing Then
                        restartAtOne = VTEquationStartsNumberScope(exact, False)
                        Set sequence = VTPureNativeNumberField(exact)
                        If Not VTEquationSequenceFieldIsCompatible(sequence, _
                           VTNativeEquationLabelName(), False, False, restartAtOne) Then
                            Set result = VTEnsurePureNativeNumberFormat(exact)
                        End If
                        Exit For
                    End If
                End If
            Next exact
        End If
    End If
    If imageFormula Then
        VTUpdateImageNumberFieldsFrom doc, changedFrom
    Else
        Set scan = doc.Range(Start:=changedFrom, End:=doc.Content.End)
        If VTCanBatchNativeNumberFields(scan) Then
            scan.Fields.Update
        Else
        For Each paragraph In scan.Paragraphs
            For Each math In paragraph.Range.OMaths
                For Each candidate In math.Range.Fields
                    If candidate.Type = wdFieldSequence Or _
                       candidate.Type = wdFieldStyleRef Then candidate.Update
                Next candidate
            Next math
        Next paragraph
        End If
        For Each candidate In doc.Fields
            If candidate.Type = wdFieldRef Then candidate.Update
        Next candidate
    End If
End Sub

Public Function VTCurrentNumberTargets( _
    ByVal doc As Document, _
    Optional ByVal referenceKind As String = "") As Collection
    Dim result As New Collection
    Dim paragraph As Paragraph
    Dim candidate As Field
    Dim source As Range
    Dim math As OMath
    Dim lastEnd As Long
    Dim paragraphHasMath As Boolean
    referenceKind = LCase$(Trim$(referenceKind))
    If Len(referenceKind) > 0 And _
       referenceKind <> VT_WORD_REFERENCE_KIND_NATIVE And _
       referenceKind <> VT_WORD_REFERENCE_KIND_IMAGE Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The Equation-reference kind is invalid: " & referenceKind
    End If
    lastEnd = -1
    For Each paragraph In doc.Paragraphs
        paragraphHasMath = False
        If referenceKind <> VT_WORD_REFERENCE_KIND_IMAGE Then
            For Each math In paragraph.Range.OMaths
                paragraphHasMath = True
                Set candidate = VTPureNativeNumberField(math.Range)
                If Not candidate Is Nothing Then
                    Set source = VTPureNativeVisibleNumberRange( _
                        math.Range, candidate)
                    If source.Start >= lastEnd Then
                        result.Add source
                        lastEnd = source.End
                    End If
                End If
            Next math
        End If
        For Each candidate In paragraph.Range.Fields
            Set source = Nothing
            If VTIsImagePlaceRef(candidate) Then
                If referenceKind <> VT_WORD_REFERENCE_KIND_NATIVE Then _
                    Set source = VTCurrentImageNumberRange(candidate)
            ElseIf candidate.Type = wdFieldSequence And _
                   referenceKind <> VT_WORD_REFERENCE_KIND_IMAGE Then
                ' Pure native fields were already collected from the paragraph's
                ' OMath objects above. Re-entering VTPureNativeNumberField through
                ' a Field.Result range is both redundant and unstable after an
                ' image-to-OMML replacement on Word for Mac. This branch is only
                ' for the legacy external helper geometry, a paragraph with no
                ' equation (a helper is never one: asking costs 5 ms an equation).
                If Not paragraphHasMath Then
                    If VTHelperParagraphOwnsNativeEquationSequence( _
                       paragraph.Range) Then
                        Set source = candidate.Result.Duplicate
                    End If
                End If
            End If
            If Not source Is Nothing Then
                If source.Start >= lastEnd Then
                    result.Add source
                    lastEnd = source.End
                End If
            End If
        Next candidate
    Next paragraph
    Set VTCurrentNumberTargets = result
End Function

' =====================================================================
' Word equation format v2 (docs/equation-format-v2.md)
'
' A numbered Word equation is written exactly as VisualTeX for Windows writes
' it: #( prefix {SEQ label [\r N] [\s L] \* ARABIC \* MERGEFORMAT} ). The app
' builds it with omml_write, the port of Windows' WordOmmlConverter
' BuildWordNativeNumberedOmml. This add-in decides the numbering context from
' the document with the rules of Windows' WordNativeNumberingAdoptionPlan and
' WordEquationNumbering, so both platforms number the same document the same
' way. No VisualTeX bookmark, variable or field stays in the document.
' =====================================================================

' --- numbered equations in the document ---------------------------------------

' The content of the #( ... ) number argument of a native numbered equation.
Private Function VTV2NumberContentRange(ByVal exact As Range) As Range
    Dim arrayFunction As OMathFunction
    Dim slotFunction As OMathFunction
    Dim marker As Range

    If exact Is Nothing Then Exit Function
    If exact.OMaths.Count <> 1 Then Exit Function
    Set marker = VTNativeEquationArrayMarkerRange(exact)
    If marker Is Nothing Then Exit Function
    For Each arrayFunction In exact.OMaths(1).Functions
        If arrayFunction.Type = wdOMathFunctionEqArray Then
            For Each slotFunction In arrayFunction.EqArray.E(1).Functions
                If slotFunction.Type = wdOMathFunctionDelim Then
                    If slotFunction.Range.Start > marker.Start Then
                        Set VTV2NumberContentRange = slotFunction.Delim.E(1).Range.Duplicate
                        Exit Function
                    End If
                End If
            Next slotFunction
        End If
    Next arrayFunction
End Function

' "SEQ Equation \r 1 \s 1 ..." -> Equation; SEQ "My Eq" -> My Eq.
Public Function VTV2SequenceName(ByVal instruction As String) As String
    Dim text As String
    Dim closing As Long
    Dim position As Long

    text = VTV2TrimSpaces(instruction)
    If UCase$(Left$(text, 4)) <> "SEQ " Then Exit Function
    text = VTV2TrimSpaces(Mid$(text, 5))
    If Left$(text, 1) = """" Then
        closing = InStr(2, text, """", vbBinaryCompare)
        If closing > 2 Then VTV2SequenceName = Mid$(text, 2, closing - 2)
        Exit Function
    End If
    For position = 1 To Len(text)
        If VTV2IsSpace(Mid$(text, position, 1)) Or Mid$(text, position, 1) = "\" Then Exit For
    Next position
    VTV2SequenceName = Left$(text, position - 1)
End Function

' The N of a "\r N" switch, or 0.
Public Function VTV2RestartValue(ByVal instruction As String) As Long
    Dim position As Long
    Dim digits As String
    position = InStr(1, instruction, "\r ", vbTextCompare)
    If position = 0 Then Exit Function
    position = position + 3
    Do While position <= Len(instruction)
        If Mid$(instruction, position, 1) <> " " Then Exit Do
        position = position + 1
    Loop
    Do While position <= Len(instruction)
        If Mid$(instruction, position, 1) < "0" Or Mid$(instruction, position, 1) > "9" Then Exit Do
        digits = digits & Mid$(instruction, position, 1)
        position = position + 1
    Loop
    If Len(digits) > 0 And Len(digits) < 9 Then VTV2RestartValue = CLng(digits)
End Function

' Windows WordNativeNumberingAdoptionPlan.TryParseLabel: "2.3-12" ->
' prefix "2.3", separator "-", ordinal 12; "7" -> ordinal 7.
Private Function VTV2ParseNumberLabel( _
    ByVal text As String, _
    ByRef prefix As String, _
    ByRef separator As String, _
    ByRef ordinal As Long) As Boolean

    Dim index As Long
    Dim character As String
    Dim ordinalText As String
    Dim candidate As String
    Dim part As Variant
    Dim parts() As String
    Dim first As String
    Dim position As Long

    prefix = ""
    separator = ""
    ordinal = 0
    index = Len(text)
    Do While index >= 1
        character = Mid$(text, index, 1)
        If character < "0" Or character > "9" Then Exit Do
        index = index - 1
    Loop
    ordinalText = Mid$(text, index + 1)
    If Len(ordinalText) = 0 Or Len(ordinalText) > 9 Then Exit Function
    If Len(ordinalText) > 1 And Left$(ordinalText, 1) = "0" Then Exit Function
    If CLng(ordinalText) < 1 Then Exit Function
    If index > 0 Then
        separator = Mid$(text, index, 1)
        If InStr(1, ".-" & ChrW(8208) & ChrW(8209) & ChrW(8210) & ChrW(8211) & ChrW(8722), _
           separator, vbBinaryCompare) = 0 Then Exit Function
        candidate = Left$(text, index - 1)
        If Len(candidate) = 0 Then Exit Function
        parts = Split(candidate, ".")
        first = parts(0)
        If Len(first) = 0 Then Exit Function
        For index = 1 To Len(first)
            character = Mid$(first, index, 1)
            If Not ((character >= "0" And character <= "9")) Then Exit For
        Next index
        If index <= Len(first) Then
            For index = 1 To Len(first)
                character = Mid$(first, index, 1)
                If Not ((character >= "A" And character <= "Z") Or (character >= "a" And character <= "z")) Then Exit Function
            Next index
        End If
        For index = 1 To UBound(parts)
            part = parts(index)
            If Len(part) = 0 Then Exit Function
            For position = 1 To Len(part)
                If Mid$(part, position, 1) < "0" Or Mid$(part, position, 1) > "9" Then Exit Function
            Next position
        Next index
        prefix = candidate
    End If
    ordinal = CLng(ordinalText)
    VTV2ParseNumberLabel = True
End Function

' Every numbered display equation with a native #(SEQ) number, in document
' order (Windows adoption plan Capture). Each item: Array(start, sequence
' name, sequence code, prefix, separator, ordinal, restart, automatic, end).
Public Function VTV2NumberedEquations( _
    ByVal doc As Document, Optional ByVal excludeStart As Long = -1) As Collection
    Dim result As New Collection
    Dim exact As Range
    Dim entry As Variant
    Dim lastStart As Long

    lastStart = -1
    For Each exact In VTNumberFieldNativeMathRoots( _
        doc, doc.Content.Start, doc.Content.End)
        If exact.Start > lastStart And exact.Start <> excludeStart Then
            entry = VTV2NumberedEquationEntry(exact)
            If Not IsEmpty(entry) Then
                result.Add entry
                lastStart = exact.Start
            End If
        End If
    Next exact
    Set VTV2NumberedEquations = result
End Function

' The numbered equations VTV2PlanInsertion needs for a number at position: every one before position, or the first one of the
' document when none is before it (entries in document order, as VTV2NumberedEquations).
Private Function VTV2NumberedEquationsAround( _
    ByVal doc As Document, ByVal position As Long, _
    Optional ByVal excludeStart As Long = -1) As Collection
    Dim result As New Collection
    Dim exact As Range
    Dim entry As Variant
    Dim lastStart As Long

    lastStart = -1
    If position > doc.Content.Start Then
        For Each exact In VTNumberFieldNativeMathRoots( _
            doc, doc.Content.Start, position)
            If exact.Start > lastStart And exact.Start <> excludeStart And _
               exact.Start < position Then
                entry = VTV2NumberedEquationEntry(exact)
                If Not IsEmpty(entry) Then
                    result.Add entry
                    lastStart = exact.Start
                End If
            End If
        Next exact
    End If
    If result.Count = 0 Then
        For Each exact In VTNumberFieldNativeMathRoots( _
            doc, position, doc.Content.End)
            If exact.Start > lastStart And exact.Start <> excludeStart Then
                entry = VTV2NumberedEquationEntry(exact)
                If Not IsEmpty(entry) Then
                    result.Add entry
                    Exit For
                End If
            End If
        Next exact
    End If
    Set VTV2NumberedEquationsAround = result
End Function

' The v2 entry of one native equation (its root Range): Array(start, sequence
' name, SEQ instruction, prefix, separator, ordinal, restart value, automatic
' heading, end), or Empty when it is not a display equation with a proven
' #(number).
Private Function VTV2NumberedEquationEntry(ByVal exact As Range) As Variant
    Dim numberField As Field
    Dim candidate As Field
    Dim number As Range
    Dim visible As String
    Dim prefix As String
    Dim separator As String
    Dim ordinal As Long
    Dim automatic As Boolean
    Dim instruction As String

    VTV2NumberedEquationEntry = Empty
    If exact Is Nothing Then Exit Function
    If exact.OMaths(1).Type <> wdOMathDisplay Then Exit Function
    Set numberField = VTPureNativeNumberField(exact)
    If numberField Is Nothing Then Exit Function
    Set number = VTV2NumberContentRange(exact)
    If number Is Nothing Then Exit Function
    For Each candidate In number.Fields
        If candidate.Type = wdFieldStyleRef Then automatic = True
    Next candidate
    number.TextRetrievalMode.IncludeFieldCodes = False
    visible = Trim$(number.Text)
    instruction = numberField.Code.Text
    If Not VTV2ParseNumberLabel(visible, prefix, separator, ordinal) Then
        ' A number WPS damaged shows a field error: still numbered, so the
        ' restyle rebuilds it (VTIsWpsDamagedNumberInstruction).
        If Not VTIsWpsDamagedNumberInstruction(instruction) Then Exit Function
        prefix = ""
        separator = ""
        ordinal = 0
    End If
    VTV2NumberedEquationEntry = Array(exact.Start, VTV2SequenceName(instruction), _
        instruction, prefix, separator, ordinal, VTV2RestartValue(instruction), _
        automatic, exact.End)
End Function

' The numbered equations VTV2PlanInsertion reads for a new number at position:
' every one of the document (VTV2NumberedEquations), or during a document
' import only the ones before position whenever there is one, because the plan
' then looks at nothing after position.
Public Function VTV2PlanEntries( _
    ByVal doc As Document, _
    ByVal position As Long, _
    ByVal excludeStart As Long) As Collection

    Dim entries As Collection
    Dim before As Collection
    Dim exact As Range
    Dim entry As Variant
    Dim lastStart As Long
    Dim searchStart As Long

    If Not VT_WORD_V2_IMPORT_CACHE_ACTIVE Or excludeStart <> -1 Then
        ' The plan reads the equations before position, or when there is none the first of the document: the ones after
        ' position are never looked at (VTV2PlanInsertion), so an insertion near the start of a long document does not
        ' read all of it.
        Set VTV2PlanEntries = VTV2NumberedEquationsAround(doc, position, excludeStart)
        Exit Function
    End If
    If VTV2ImportCacheUsable(doc, position) Then
        If Not VT_WORD_V2_IMPORT_ENTRIES Is Nothing Then
            If VT_WORD_V2_IMPORT_ENTRIES.Count > 0 Then Set before = VT_WORD_V2_IMPORT_ENTRIES
        End If
    End If
    If Not before Is Nothing Then
        ' Only equations after the last known one can be new.
        entry = before(before.Count)
        lastStart = entry(0)
        searchStart = entry(8)
        For Each exact In VTNumberFieldNativeMathRoots(doc, searchStart, position)
            If exact.Start > lastStart And exact.Start < position Then
                entry = VTV2NumberedEquationEntry(exact)
                If Not IsEmpty(entry) Then
                    before.Add entry
                    lastStart = exact.Start
                End If
            End If
        Next exact
    Else
        Set entries = VTV2NumberedEquations(doc)
        Set before = New Collection
        For Each entry In entries
            If entry(0) < position Then before.Add entry
        Next entry
        If before.Count = 0 Then
            ' Nothing before position: the plan uses the first number of
            ' the whole document.
            Set VT_WORD_V2_IMPORT_ENTRIES = before
            VTV2RememberImportPosition doc, position
            Set VTV2PlanEntries = entries
            Exit Function
        End If
    End If
    Set VT_WORD_V2_IMPORT_ENTRIES = before
    VTV2RememberImportPosition doc, position
    Set VTV2PlanEntries = before
End Function

' The number of an equation written by an editor commit. An existing v2 number
' is kept exactly as it is (Windows keeps a numbered equation's number host
' verbatim when its body is edited); anything else, including an older Mac
' number, is numbered as a new insertion (Windows PlanInsertion) that leaves
' the edited equation itself out of the series.
Public Sub VTV2PlanCommitNumber( _
    ByVal doc As Document, _
    ByVal targetRange As Range, _
    ByVal targetIsNative As Boolean, _
    ByRef sequenceName As String, _
    ByRef headingLevel As Long, _
    ByRef separator As String, _
    ByRef literalPrefix As String, _
    ByRef automatic As Boolean, _
    ByRef restart As Long, _
    ByRef initial As Long, _
    ByRef transferEquation As Range)

    Dim exact As Range
    Dim targetEntry As Variant
    Dim excludeStart As Long
    Dim instruction As String
    Dim levelPosition As Long

    excludeStart = -1
    Set transferEquation = Nothing
    VTWordPerformanceMark "v2-plan-start"
    If targetIsNative Then
        If targetRange.OMaths.Count >= 1 Then
            Set exact = VTNativeMathRoot(targetRange.OMaths(1)).Range.Duplicate
            excludeStart = exact.Start
            targetEntry = VTV2NumberedEquationEntry(exact)
            If Not IsEmpty(targetEntry) Then
                instruction = targetEntry(2)
                headingLevel = 0
                levelPosition = InStr(1, instruction, "\s ", vbTextCompare)
                If levelPosition > 0 Then _
                    headingLevel = CLng(Val(Mid$(instruction, levelPosition + 3, 2)))
                separator = targetEntry(4)
                automatic = targetEntry(7)
                literalPrefix = ""
                If headingLevel > 0 And Not automatic Then literalPrefix = targetEntry(3)
                restart = targetEntry(6)
                If (headingLevel = 0 Or separator = "." Or separator = "-") And _
                   VTV2NumberHasStyle(exact, targetEntry, headingLevel, separator, _
                       literalPrefix, automatic, restart) Then
                    sequenceName = targetEntry(1)
                    initial = targetEntry(5)
                    Exit Sub
                End If
            End If
        End If
    End If
    VTV2PlanInsertion doc, targetRange.Start, sequenceName, headingLevel, separator, _
        literalPrefix, automatic, restart, initial, transferEquation, excludeStart
    VTWordPerformanceMark "v2-plan-complete"
End Sub

' --- the app builds the number (omml_write) ----------------------------------

Public Function VTV2NumberItemManifest( _
    ByVal index As Long, _
    ByVal sequenceName As String, _
    ByVal headingLevel As Long, _
    ByVal separator As String, _
    ByVal literalPrefix As String, _
    ByVal automatic As Boolean, _
    ByVal restart As Long, _
    ByVal initial As Long) As String

    Dim key As String
    key = "item" & CStr(index)
    VTV2NumberItemManifest = _
        key & "SequenceName=" & VTBase64UrlEncodeUtf8(sequenceName) & vbLf & _
        key & "HeadingLevel=" & CStr(headingLevel) & vbLf & _
        key & "Separator=" & VTBase64UrlEncodeUtf8(separator) & vbLf & _
        IIf(headingLevel > 0 And Not automatic, _
            key & "LiteralPrefix=" & VTBase64UrlEncodeUtf8(literalPrefix) & vbLf, "") & _
        IIf(restart > 0, key & "Restart=" & CStr(restart) & vbLf, "") & _
        key & "Initial=" & CStr(IIf(initial > 0, initial, 1)) & vbLf
End Function

' Sends one nativeNumber request (several numbered equations; one body or
' VisualTeX's x placeholder) and returns the .docx the app wrote.
Public Function VTV2RequestNumberDocx( _
    ByVal sessionId As String, _
    ByVal doc As Document, _
    ByVal ommlBase64 As String, _
    ByVal itemCount As Long, _
    ByVal items As String) As String

    Dim manifest As String
    Dim timingDetail As String
    Dim encodedResult As String
    Dim separatorPosition As Long
    Dim result As Object
    Dim docxPath As String
    Dim resultText As String
    Dim fastTrace As String
    Dim requestStep As String
    Dim failedNumber As Long
    Dim failedSource As String
    Dim failedDescription As String

    On Error GoTo RequestFailed
    requestStep = "font"
    VTApplyChosenOmmlFontToDocument doc
    VTWordPerformanceMark "v2-font-applied"
    requestStep = "manifest"
    manifest = "protocolVersion=" & CStr(VT_PROTOCOL_VERSION) & vbLf & _
        "sessionId=" & sessionId & vbLf & _
        "sourceKind=nativeNumber" & vbLf & _
        "sourceDocumentId=" & VTBase64UrlEncodeUtf8(doc.FullName) & vbLf & _
        "bookmarkName=-" & vbLf & _
        "itemCount=" & CStr(itemCount) & vbLf & _
        "mathFont=" & VTBase64UrlEncodeUtf8(doc.OMathFontName) & vbLf & _
        IIf(Len(ommlBase64) > 0, "ommlBase64=" & ommlBase64 & vbLf, "") & items
    VTWordPerformanceMark "v2-manifest-built"
    ' The resident app builds the package at once; the helper script's process answers in half a second.
    requestStep = "resident"
    resultText = VTTryFastConvertOmmlBatch(sessionId, manifest, fastTrace)
    VTWordPerformanceNote "v2-resident-trace", fastTrace
    If Len(resultText) > 0 Then
        VTWordPerformanceMark "v2-app-converted-fast"
    Else
        VTWordPerformanceMark "v2-resident-declined"
        requestStep = "direct-write"
        If VTTryWriteOmmlBatchManifestDirect(sessionId, manifest) Then
            requestStep = "script-convert"
            timingDetail = VTConvertOmmlBatch(VT_WORD_HOST, sessionId)
        Else
            requestStep = "script-convert-argument"
            timingDetail = VTConvertOmmlBatch(VT_WORD_HOST, sessionId, manifest)
        End If
        VTWordPerformanceMark "v2-app-converted"
        VTWordPerformanceNote "v2-app-timing", Left$(timingDetail, InStr(1, timingDetail & "|", "|", vbBinaryCompare) - 1)
        separatorPosition = InStrRev(timingDetail, "|", -1, vbBinaryCompare)
        If separatorPosition <= 1 Or separatorPosition >= Len(timingDetail) Then
            Err.Raise vbObjectError + 7611, "VisualTeX", _
                "VisualTeX did not return the Word equation number."
        End If
        encodedResult = Mid$(timingDetail, separatorPosition + 1)
        resultText = VTBase64UrlDecodeUtf8(encodedResult)
    End If
    requestStep = "parse"
    Set result = VTParseDocumentImportManifestText(resultText, sessionId)
    VTRequireDispatchValue result, "docxPath"
    If CStr(result("operation")) <> "nativeNumber" Or _
       CLng(result("itemCount")) <> itemCount Then
        Err.Raise vbObjectError + 7611, "VisualTeX", _
            "VisualTeX returned a Word equation number for another request."
    End If
    docxPath = CStr(result("docxPath"))
    requestStep = "package"
    VTValidateAbsoluteVisualTeXPath docxPath
    If Not VTPathFileExists(docxPath) Then
        Err.Raise vbObjectError + 7611, "VisualTeX", _
            "The Word equation number package is missing."
    End If
    VTV2RequestNumberDocx = docxPath
    Exit Function

RequestFailed:
    ' the same error goes on to the caller; the trace says which step it came from
    failedNumber = Err.Number
    failedSource = Err.Source
    failedDescription = Err.Description
    VTWordPerformanceNote "v2-request-failed", requestStep & ": " & CStr(failedNumber) & " " & failedDescription
    VTWordPerformanceNote "v2-resident-trace", fastTrace
    On Error GoTo 0
    Err.Raise failedNumber, failedSource, failedDescription
End Function

' --- after a numbered equation is inserted -----------------------------------

' Moves the \r switch of the series' former first number to the new number
' (Windows ChangeKind.TransferReset) and lets Word recount: STYLEREF before
' SEQ inside every equation from the new one on, then the references.
Public Sub VTV2FinishInsertedNumber( _
    ByVal doc As Document, _
    ByVal exact As Range, _
    ByVal transferEquation As Range)

    VTV2TransferRestart transferEquation
    VTV2UpdateNumbersFrom doc, exact.Start
End Sub

' Removes the \r switch from the number of transferEquation, the former first
' number of a series that a new number now starts.
Public Sub VTV2TransferRestart(ByVal transferEquation As Range)
    Dim owner As Range
    Dim candidate As Field
    Dim code As Range
    Dim instruction As String
    Dim switchStart As Long
    Dim switchText As String

    If transferEquation Is Nothing Then Exit Sub
    If transferEquation.OMaths.Count < 1 Then Exit Sub
    Set owner = VTNativeMathRoot(transferEquation.OMaths(1)).Range.Duplicate
    Set candidate = VTPureNativeNumberField(owner)
    If candidate Is Nothing Then Exit Sub
    instruction = candidate.Code.Text
    switchStart = InStr(1, instruction, "\r ", vbTextCompare)
    If switchStart > 0 And VTV2RestartValue(instruction) > 0 Then
        ' Windows replaces exactly the switch with one space and never
        ' rewrites the rest of the instruction.
        switchText = Mid$(instruction, switchStart, _
            InStr(switchStart + 3, instruction & " ", " ", vbBinaryCompare) - switchStart)
        If Len(switchText) <= 3 Then switchText = "\r " & CStr(VTV2RestartValue(instruction))
        Set code = candidate.Code.Duplicate
        code.SetRange code.Start + switchStart - 1, _
            code.Start + switchStart - 1 + Len(switchText)
        If code.Text = switchText Then code.Text = " "
    End If
End Sub

' STYLEREF before SEQ in every numbered equation from position on (Windows
' WordNativeOmmlFieldRefresh), then every cross-reference.
Public Sub VTV2UpdateNumbersFrom( _
    ByVal doc As Document, ByVal fromPosition As Long, _
    Optional ByVal updateReferences As Boolean = True)
    Dim owner As Range
    Dim candidate As Field
    Dim lastStart As Long
    Dim searchStart As Long

    ' Only equations holding STYLEREF/SEQ fields have anything to update, and
    ' an equation ending at or after fromPosition starts no earlier than the
    ' paragraph that contains fromPosition.
    searchStart = fromPosition
    If searchStart < doc.Content.Start Then searchStart = doc.Content.Start
    If searchStart > doc.Content.End Then searchStart = doc.Content.End
    searchStart = doc.Range(Start:=searchStart, End:=searchStart) _
        .Paragraphs(1).Range.Start
    lastStart = -1
    Dim ownerRoots As Collection
    Dim updatedOwners As Long
    Dim updatedReferences As Long
    VTWordPerformanceMark "renumber-start"
    If VTV2BulkRenumber(doc, searchStart, updateReferences) Then
        VTWordPerformanceMark "renumber-bulk-complete"
        Exit Sub
    End If
    Set ownerRoots = VTNumberFieldNativeMathRoots( _
        doc, searchStart, doc.Content.End, True)
    VTWordPerformanceMark "renumber-roots-listed"
    VTWordPerformanceNote "renumber-roots", CStr(ownerRoots.Count)
    For Each owner In ownerRoots
        If owner.End >= fromPosition And owner.Start > lastStart Then
            lastStart = owner.Start
            updatedOwners = updatedOwners + 1
            For Each candidate In owner.Fields
                If candidate.Type = wdFieldStyleRef Then candidate.Update
            Next candidate
            For Each candidate In owner.Fields
                If candidate.Type = wdFieldSequence Then candidate.Update
            Next candidate
        End If
    Next owner
    VTWordPerformanceMark "renumber-fields-updated"
    VTWordPerformanceNote "renumber-updated-equations", CStr(updatedOwners)
    If Not updateReferences Then Exit Sub
    For Each candidate In doc.Fields
        If candidate.Type = wdFieldRef Then
            candidate.Update
            updatedReferences = updatedReferences + 1
        End If
    Next candidate
    VTWordPerformanceMark "renumber-references-updated"
    VTWordPerformanceNote "renumber-updated-references", CStr(updatedReferences)
End Sub

' --- renumbering in bulk ---------------------------------------------------------------------------------------------
' Word for Mac (measured, 125 numbered equations): Field.Update takes about 6 ms a call, and Fields.Update on a Range 0.3 ms a
' field, so the numbers from the new one on are updated with one call over the run of numbering fields instead of one call
' for each, which took 8 ms an equation (a start-of-document insertion in 400 equations: 3.3 s).

' The fields VTV2UpdateNumbersFrom updates: the SEQ and STYLEREF of a native equation (whatever its label), the literal
' chapter prefix QUOTE of its number, and the references. A SEQ or STYLEREF outside an equation (a figure caption, an image
' number), a QUOTE that is calculated or the user's, and every other kind of field keep their own chain.
Private Function VTV2FieldIsNumbering( _
    ByVal candidate As Field, ByVal fieldType As Long) As Boolean

    Select Case fieldType
        Case wdFieldRef
            VTV2FieldIsNumbering = True
        Case wdFieldSequence, wdFieldStyleRef
            VTV2FieldIsNumbering = Not VTFieldNativeMathRoot(candidate) Is Nothing
        Case wdFieldQuote
            VTV2FieldIsNumbering = (StrComp(Trim$(candidate.Code.Text), _
                VTNativePrefixInstruction(candidate.Result.Text), vbTextCompare) = 0)
    End Select
End Function

' The other fields an update in the same call may pass over: ones an update leaves as they are (an add-in's ADDIN fields,
' hyperlinks). Any other (a table of contents, a date, an include, a picture or object link ...) ends the run.
Private Function VTV2FieldTypeUpdatesHarmlessly(ByVal fieldType As Long) As Boolean
    Select Case fieldType
        Case wdFieldAddin, wdFieldHyperlink
            VTV2FieldTypeUpdatesHarmlessly = True
    End Select
End Function

' Updates the numbering fields from fromPosition to the end of the document: each run of them (with only harmless fields
' between) in one Fields.Update. Returns the number of calls; sawReference tells whether a reference was among them.
Private Function VTV2UpdateNumberingFields( _
    ByVal doc As Document, ByVal fromPosition As Long, _
    ByRef sawReference As Boolean) As Long

    Dim candidate As Field
    Dim fieldType As Long
    Dim segmentStart As Long
    Dim segmentEnd As Long
    Dim inSegment As Boolean
    Dim calls As Long

    For Each candidate In doc.Range( _
        Start:=fromPosition, End:=doc.Content.End).Fields
        fieldType = candidate.Type
        If VTV2FieldIsNumbering(candidate, fieldType) Then
            If fieldType = wdFieldRef Then sawReference = True
            If Not inSegment Then
                segmentStart = VTEquationFieldStart(candidate)
                inSegment = True
            End If
            segmentEnd = VTEquationFieldEnd(candidate)
        ElseIf Not VTV2FieldTypeUpdatesHarmlessly(fieldType) Then
            If inSegment Then
                doc.Range(Start:=segmentStart, End:=segmentEnd).Fields.Update
                calls = calls + 1
                inSegment = False
            End If
        End If
    Next candidate
    If inSegment Then
        doc.Range(Start:=segmentStart, End:=segmentEnd).Fields.Update
        calls = calls + 1
    End If
    VTV2UpdateNumberingFields = calls
End Function

' The renumbering of VTV2UpdateNumbersFrom in bulk; False when Word refuses somewhere (the caller then does it field by field).
Private Function VTV2BulkRenumber( _
    ByVal doc As Document, ByVal searchStart As Long, _
    ByVal updateReferences As Boolean) As Boolean

    Dim candidate As Field
    Dim calls As Long
    Dim sawReference As Boolean
    Dim referencesBefore As Boolean

    On Error GoTo NotBulk
    calls = VTV2UpdateNumberingFields(doc, searchStart, sawReference)
    VTWordPerformanceMark "renumber-numbers-updated"
    VTWordPerformanceNote "renumber-update-calls", CStr(calls)
    If updateReferences Then
        ' The references of the document: the ones before the new number read the numbers just updated, and the ones after
        ' it that stand ahead of the equation they point to were updated with a stale number. A Field.Update takes
        ' about 10 ms for a reference in Word for Mac, so they are updated in bulk as well, when there are any.
        If Not sawReference And searchStart > doc.Content.Start Then
            For Each candidate In doc.Range( _
                Start:=doc.Content.Start, End:=searchStart).Fields
                If candidate.Type = wdFieldRef Then
                    referencesBefore = True
                    Exit For
                End If
            Next candidate
        End If
        If sawReference Or referencesBefore Then
            calls = calls + VTV2UpdateNumberingFields(doc, doc.Content.Start, sawReference)
        End If
        VTWordPerformanceMark "renumber-references-updated"
        VTWordPerformanceNote "renumber-update-calls", CStr(calls)
    End If
    VTV2BulkRenumber = True
    Exit Function

NotBulk:
    Err.Clear
End Function

' --- restyle and renumber (Windows PlanFormat) ---------------------------------

' The tokens of a SEQ instruction: name, \r value, \s level, and whether the
' ARABIC and MERGEFORMAT switches of the shared format are present.
Private Function VTV2SequenceMatches( _
    ByVal instruction As String, _
    ByVal sequenceName As String, _
    ByVal restart As Long, _
    ByVal headingLevel As Long) As Boolean

    Dim levelText As String
    Dim position As Long

    If StrComp(VTV2SequenceName(instruction), sequenceName, vbTextCompare) <> 0 Then Exit Function
    If VTV2RestartValue(instruction) <> restart Then Exit Function
    position = InStr(1, instruction, "\s ", vbTextCompare)
    If position > 0 Then levelText = Trim$(Mid$(instruction, position + 3, 2))
    If headingLevel = 0 Then
        If position > 0 Then Exit Function
    ElseIf Left$(levelText, 1) <> CStr(headingLevel) Then
        Exit Function
    End If
    If InStr(1, instruction, "\* ARABIC", vbTextCompare) = 0 Then Exit Function
    If InStr(1, instruction, "\* MERGEFORMAT", vbTextCompare) = 0 Then Exit Function
    VTV2SequenceMatches = True
End Function

' True when the equation's number already has exactly the requested form.
Private Function VTV2NumberHasStyle( _
    ByVal exact As Range, _
    ByVal entry As Variant, _
    ByVal headingLevel As Long, _
    ByVal separator As String, _
    ByVal literalPrefix As String, _
    ByVal automatic As Boolean, _
    ByVal restart As Long) As Boolean

    Dim number As Range
    Dim candidate As Field
    Dim styleRefs As Long
    Dim others As Long

    If Not VTV2SequenceMatches(entry(2), entry(1), restart, headingLevel) Then Exit Function
    Set number = VTV2NumberContentRange(exact)
    If number Is Nothing Then Exit Function
    For Each candidate In number.Fields
        Select Case candidate.Type
            Case wdFieldSequence
            Case wdFieldStyleRef
                styleRefs = styleRefs + 1
                If InStr(1, candidate.Code.Text, "STYLEREF " & CStr(headingLevel) & " ", vbTextCompare) = 0 Then Exit Function
            Case Else
                others = others + 1
        End Select
    Next candidate
    If others > 0 Then Exit Function
    If headingLevel = 0 Then
        VTV2NumberHasStyle = (styleRefs = 0 And Len(entry(3)) = 0)
    ElseIf automatic Then
        VTV2NumberHasStyle = (styleRefs = 1 And entry(4) = separator)
    Else
        VTV2NumberHasStyle = (styleRefs = 0 And entry(3) = literalPrefix And entry(4) = separator)
    End If
End Function

' Rewrites the number of every numbered equation whose number is not already
' in the document's numbering style (Windows PlanFormat): its own label, the
' style's prefix and separator, \s L, and \r 1 on the first number of each
' chapter/section (or of the document). The app builds all new numbers in one
' package; each is copied into its equation's number slot, so equation bodies
' and their run formatting are untouched. Returns how many were rewritten.
Public Function VTV2RestyleNumbers( _
    ByVal doc As Document, Optional ByVal dryRun As Boolean = False) As Long
    Dim entries As Collection
    Dim headingLevel As Long
    Dim separator As String
    Dim index As Long
    Dim scopeStart As Long
    Dim automatic As Boolean
    Dim numberText As String
    Dim counterKeys() As String
    Dim counterValues() As Long
    Dim counterCount As Long
    Dim key As String
    Dim slot As Long
    Dim ordinal As Long
    Dim restart As Long
    Dim literalPrefix As String
    Dim changed() As Long
    Dim changedCount As Long
    Dim items As String
    Dim sessionId As String
    Dim docxPath As String
    Dim staging As Document
    Dim exact As Range
    Dim target As Range
    Dim source As Range
    Dim names As Collection
    Dim name As Variant
    Dim firstChange As Long
    Dim headings As Collection

    Set entries = VTV2NumberedEquations(doc)
    If entries.Count = 0 Then Exit Function
    VTV2DocumentFormat doc, headingLevel, separator
    Set headings = VTV2CollectHeadings(doc, headingLevel)
    ReDim counterKeys(1 To entries.Count)
    ReDim counterValues(1 To entries.Count)
    ReDim changed(1 To entries.Count)
    firstChange = -1
    For index = 1 To entries.Count
        literalPrefix = ""
        automatic = False
        scopeStart = -1
        If headingLevel > 0 Then
            numberText = VTV2HeadingScope(headings, entries(index)(0), headingLevel, scopeStart, automatic)
            If scopeStart < 0 Then automatic = False
            If Not automatic Then literalPrefix = numberText
        End If
        key = LCase$(entries(index)(1)) & "|" & _
            IIf(headingLevel = 0, "continuous", CStr(scopeStart))
        For slot = 1 To counterCount
            If counterKeys(slot) = key Then Exit For
        Next slot
        If slot > counterCount Then
            counterCount = counterCount + 1
            counterKeys(counterCount) = key
            slot = counterCount
        End If
        counterValues(slot) = counterValues(slot) + 1
        ordinal = counterValues(slot)
        restart = IIf(ordinal = 1, 1, 0)
        Set exact = doc.Range(entries(index)(0), entries(index)(8))
        If Not VTV2NumberHasStyle(exact, entries(index), headingLevel, separator, _
            literalPrefix, automatic, restart) Then
            changedCount = changedCount + 1
            changed(changedCount) = index
            items = items & VTV2NumberItemManifest(changedCount, entries(index)(1), _
                headingLevel, separator, literalPrefix, automatic, restart, ordinal)
            If firstChange < 0 Then firstChange = entries(index)(0)
        End If
    Next index
    If dryRun Then
        VTV2RestyleNumbers = changedCount
        Exit Function
    End If
    If changedCount = 0 Then
        VTV2UpdateNumbersFrom doc, 0
        Exit Function
    End If

    sessionId = VTNewUuidV4()
    On Error GoTo Failed
    docxPath = VTV2RequestNumberDocx(sessionId, doc, "", changedCount, items)
    Set staging = Documents.Open(FileName:=docxPath, ConfirmConversions:=False, _
        ReadOnly:=True, AddToRecentFiles:=False, Visible:=False)
    If staging.OMaths.Count <> changedCount Then
        Err.Raise vbObjectError + 7612, "VisualTeX", _
            "Word did not open every prepared equation number."
    End If
    doc.Activate
    ' Last to first: copying a number changes only text after its own start.
    For index = changedCount To 1 Step -1
        Set exact = doc.Range(entries(changed(index))(0), entries(changed(index))(8))
        Set exact = VTNativeMathRoot(exact.OMaths(1)).Range.Duplicate
        Set names = VTCaptureNativeNumberBookmarks(exact)
        For Each name In names
            doc.Bookmarks(CStr(name)).Delete
        Next name
        Set target = VTV2NumberContentRange(exact)
        Set source = VTV2NumberContentRange( _
            VTNativeMathRoot(staging.OMaths(index)).Range.Duplicate)
        If target Is Nothing Or source Is Nothing Then
            Err.Raise vbObjectError + 7612, "VisualTeX", _
                "An equation number slot could not be resolved."
        End If
        target.FormattedText = source.FormattedText
        Set exact = doc.Range(entries(changed(index))(0), entries(changed(index))(0))
        Set exact = VTNativeMathRoot(exact.Paragraphs(1).Range.OMaths(1)).Range.Duplicate
        Set target = VTV2NumberContentRange(exact)
        For Each name In names
            doc.Bookmarks.Add Name:=CStr(name), Range:=target
        Next name
    Next index
    staging.Close SaveChanges:=wdDoNotSaveChanges
    Set staging = Nothing
    doc.Activate
    VTDeleteSessionFiles sessionId
    VTV2UpdateNumbersFrom doc, firstChange
    VTV2RestyleNumbers = changedCount
    Exit Function

Failed:
    Dim failureNumber As Long
    Dim failureDescription As String
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Not staging Is Nothing Then staging.Close SaveChanges:=wdDoNotSaveChanges
    doc.Activate
    VTDeleteSessionFiles sessionId
    On Error GoTo 0
    Err.Raise failureNumber, "VisualTeX equation numbers", failureDescription
End Function

' Word cross-reference bookmarks (_Ref...) of a number that is rebuilt are put
' back on the new number content, so body REF fields keep their target.
Public Sub VTV2RestoreNumberBookmarks( _
    ByVal equationRange As Range, ByVal names As Collection)

    Dim exact As Range
    Dim number As Range
    Dim name As Variant

    If names Is Nothing Or equationRange Is Nothing Then Exit Sub
    If names.Count = 0 Or equationRange.OMaths.Count < 1 Then Exit Sub
    Set exact = VTNativeMathRoot(equationRange.OMaths(1)).Range.Duplicate
    Set number = VTV2NumberContentRange(exact)
    If number Is Nothing Then Exit Sub
    For Each name In names
        If exact.Document.Bookmarks.Exists(CStr(name)) Then _
            exact.Document.Bookmarks(CStr(name)).Delete
        exact.Document.Bookmarks.Add Name:=CStr(name), Range:=number
    Next name
End Sub

Public Function VTEnsurePureNativeNumberFormat( _
    ByVal equationRange As Range, _
    Optional ByVal headingPrefixOverride As String = "", _
    Optional ByVal headingPrefixResolved As Boolean = False, _
    Optional ByVal restartAtOneOverride As Boolean = False, _
    Optional ByVal restartResolved As Boolean = False) As Range

    Dim documentObject As Document
    Dim exact As Range
    Dim marker As Range
    Dim numberField As Field
    Dim headingField As Field
    Dim prefixField As Field
    Dim candidate As Field
    Dim placeholder As Range
    Dim editRange As Range
    Dim preservedBookmarks As Collection
    Dim desiredLevel As Long
    Dim headingCount As Long
    Dim prefixCount As Long
    Dim numberInstruction As String
    Dim headingPrefix As String
    Dim expectedPrefixText As String
    Dim expectedRestartAtOne As Boolean
    Dim structureMatches As Boolean
    Dim operationStage As String
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim slotStart As Long
    Dim backupDocument As Document
    Dim backupRange As Range
    Dim restoreMath As OMath
    Dim originalStart As Long
    Dim originalText As String
    Dim xmlStart As Long
    Dim xmlEnd As Long
    Dim mutationStarted As Boolean
    Dim rollbackFailure As String

    On Error GoTo FormatFailed
    operationStage = "resolve"
    If equationRange Is Nothing Or equationRange.OMaths.Count <> 1 Then Exit Function
    Set documentObject = equationRange.Document
    Set exact = VTNativeMathRoot(equationRange.OMaths(1)).Range.Duplicate
    Set marker = VTNativeEquationArrayMarkerRange(exact)
    If marker Is Nothing Then Exit Function
    Set numberField = VTPureNativeNumberField(exact)
    If numberField Is Nothing Then Exit Function
    VTPinEquationNumbering documentObject

    desiredLevel = VTNativeHeadingLevelAtRange(exact)
    If desiredLevel > 0 Then
        If headingPrefixResolved Then
            headingPrefix = headingPrefixOverride
        Else
            headingPrefix = VTEquationHeadingPrefix( _
                exact, VTEquationNumberingMode(documentObject))
        End If
        expectedPrefixText = headingPrefix & _
            VTEquationNumberingDisplaySeparator(documentObject)
        If restartResolved Then
            expectedRestartAtOne = restartAtOneOverride
        Else
            expectedRestartAtOne = VTEquationStartsNumberScope( _
                exact, False)
        End If
    End If
    numberInstruction = VTNativeSequenceInstruction( _
        documentObject, False, False, expectedRestartAtOne)
    For Each candidate In exact.Fields
        If candidate.Type = wdFieldStyleRef And _
           VTEquationFieldStart(candidate) >= marker.End Then
            headingCount = headingCount + 1
            If headingField Is Nothing Then
                Set headingField = candidate
            ElseIf VTEquationFieldStart(candidate) < _
                   VTEquationFieldStart(headingField) Then
                Set headingField = candidate
            End If
        ElseIf candidate.Type = wdFieldQuote And _
               VTEquationFieldStart(candidate) >= marker.End And _
               candidate.Result.Start <= numberField.Result.Start Then
            prefixCount = prefixCount + 1
            Set prefixField = candidate
        End If
    Next candidate

    structureMatches = VTEquationSequenceFieldIsCompatible( _
        numberField, VTNativeEquationLabelName(), False, False, _
        expectedRestartAtOne)
    structureMatches = structureMatches And headingCount = 0
    If structureMatches Then
        If Len(expectedPrefixText) > 0 Then
            structureMatches = prefixCount = 1
            If structureMatches Then
                structureMatches = StrComp( _
                    prefixField.Result.Text, expectedPrefixText, _
                    vbBinaryCompare) = 0
            End If
        Else
            structureMatches = prefixCount = 0
        End If
    End If
    If structureMatches Then
        numberField.Update
        Set VTEnsurePureNativeNumberFormat = exact.Duplicate
        Exit Function
    End If

    Set preservedBookmarks = VTCaptureNativeNumberBookmarks(exact)
    originalStart = exact.Start
    originalText = VTXmlTextContent(VTNativeMathXmlBody(exact.WordOpenXML, xmlStart, xmlEnd))
    Set backupDocument = Documents.Add(Visible:=False)
    Set backupRange = backupDocument.Content.Duplicate
    backupRange.Collapse wdCollapseStart
    backupRange.FormattedText = exact.FormattedText
    Set backupRange = VTNativeMathRoot(backupDocument.OMaths(1)).Range.Duplicate
    documentObject.Activate
    slotStart = VTEquationFieldStart(numberField)
    If Not headingField Is Nothing Then
        slotStart = VTEquationFieldStart(headingField)
    End If

    ' Word for Mac reparses an in-place Code.Text assignment inside OMath and
    ' can reinterpret a localized SEQ identifier as a Bookmark. Rebuild only the
    ' owned right-hand number slot through the same native Fields.Add path used
    ' for initial insertion. The formula XML before the # marker is untouched.
    operationStage = "unlink-heading-prefix"
    mutationStarted = True
    If prefixCount > 1 Or headingCount > 1 Then
        Err.Raise vbObjectError + 7606, "VisualTeX", _
            "The native Equation number contains ambiguous prefix fields."
    End If
    If Not prefixField Is Nothing Then
        prefixField.Result.Text = ""
        prefixField.Unlink
        Set exact = VTNativeMathRoot(equationRange.Paragraphs(1).Range.OMaths(1)).Range.Duplicate
        Set numberField = VTPureNativeNumberField(exact)
        If numberField Is Nothing Then
            Err.Raise vbObjectError + 7606, "VisualTeX", _
                "The native Equation sequence disappeared with its static prefix."
        End If
    ElseIf Not headingField Is Nothing Then
        headingField.Result.Text = ""
        headingField.Unlink
        Set exact = VTNativeMathRoot(equationRange.Paragraphs(1).Range.OMaths(1)).Range.Duplicate
        Set numberField = VTPureNativeNumberField(exact)
        If numberField Is Nothing Then
            Err.Raise vbObjectError + 7606, "VisualTeX", _
                "The native Equation sequence disappeared with its heading prefix."
        End If
        Set editRange = documentObject.Range( _
            slotStart, VTEquationFieldStart(numberField))
        If editRange.End > editRange.Start Then editRange.Delete
        Set exact = VTNativeMathRoot(equationRange.Paragraphs(1).Range.OMaths(1)).Range.Duplicate
        Set numberField = VTPureNativeNumberField(exact)
    End If
    operationStage = "unlink-sequence"
    numberField.Result.Text = VT_WORD_NATIVE_NUMBER_CACHE_PLACEHOLDER
    numberField.Unlink

    Set exact = VTNativeMathRoot(equationRange.Paragraphs(1).Range.OMaths(1)).Range.Duplicate
    Set placeholder = VTNativeEquationArrayPlaceholderRange(exact)
    If placeholder Is Nothing Then
        Err.Raise vbObjectError + 7606, "VisualTeX", _
            "The native Equation number slot could not be staged for rebuilding."
    End If
    operationStage = "insert-sequence"
    Set numberField = documentObject.Fields.Add( _
        Range:=placeholder, Type:=wdFieldEmpty, _
        Text:=numberInstruction, PreserveFormatting:=True)
    If Len(expectedPrefixText) > 0 Then
        Set editRange = documentObject.Range( _
            VTEquationFieldStart(numberField), _
            VTEquationFieldStart(numberField))
        Set prefixField = VTInsertFieldBefore( _
            editRange, VTNativePrefixInstruction(expectedPrefixText), True)
        prefixField.Update
    End If

    operationStage = "finalize"
    Set exact = VTNativeMathRoot(numberField.Result.Paragraphs(1).Range.OMaths(1)).Range.Duplicate
    numberField.Update
    VTRestoreNativeNumberBookmarks exact, preservedBookmarks
    exact.OMaths(1).Type = wdOMathDisplay
    exact.OMaths(1).Justification = wdOMathJcCenterGroup
    ' This path owns only the right-hand number slot of an existing native
    ' equation. Reapplying the Caption paragraph style here clears direct font
    ' formatting from the formula body on Word for Mac (for example 15 pt is
    ' reset to the Caption style's 10 pt). Paragraph styling belongs to initial
    ' equation creation; a numbering-format switch must leave it untouched.
    backupDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set backupDocument = Nothing
    documentObject.Activate
    Set VTEnsurePureNativeNumberFormat = exact.Duplicate
    Exit Function

FormatFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If mutationStarted And Not backupRange Is Nothing Then
        documentObject.Activate
        Set restoreMath = VTNativeMathNearStart(documentObject, originalStart, 8)
        Set exact = restoreMath.Range.Duplicate
        exact.FormattedText = backupRange.FormattedText
        Set exact = VTNativeMathRootRange(exact)
        VTRestoreNativeNumberBookmarks exact, preservedBookmarks
        If VTXmlTextContent(VTNativeMathXmlBody(exact.WordOpenXML, xmlStart, xmlEnd)) <> originalText Then _
            rollbackFailure = "; Word did not verify the original number after rollback."
        If Err.Number <> 0 Then rollbackFailure = "; number rollback failed: " & Err.Description
    End If
    If Not backupDocument Is Nothing Then backupDocument.Close SaveChanges:=wdDoNotSaveChanges
    documentObject.Activate
    On Error GoTo 0
    Err.Raise failureNumber, "VisualTeX native number format", _
        operationStage & ": " & failureDescription & rollbackFailure
End Function

Public Sub VTRestoreNativeNumberBookmarks(ByVal equation As Range, ByVal names As Collection)
    Dim number As Field
    Dim source As Range
    Dim item As Variant
    If names Is Nothing Then Exit Sub
    Set number = VTPureNativeNumberField(equation)
    If number Is Nothing Then Exit Sub
    Set source = VTPureNativeVisibleNumberRange(equation, number)
    If source Is Nothing Then Exit Sub
    For Each item In names
        equation.Document.Bookmarks.Add Name:=CStr(item), Range:=source
    Next item
End Sub

Public Function VTCaptureNativeNumberBookmarks(ByVal equation As Range) As Collection
    Dim names As New Collection
    Dim number As Field
    Dim mark As Bookmark
    Dim hidden As Boolean
    Dim doc As Document
    Set doc = equation.Document
    Set number = VTPureNativeNumberField(equation)
    If Not number Is Nothing Then
        hidden = doc.Bookmarks.ShowHidden
        doc.Bookmarks.ShowHidden = True
        For Each mark In doc.Bookmarks
            If Left$(mark.Name, 4) = "_Ref" And _
               mark.Range.Start >= equation.Start And mark.Range.End <= equation.End Then
                names.Add mark.Name
            End If
        Next mark
        doc.Bookmarks.ShowHidden = hidden
    End If
    Set VTCaptureNativeNumberBookmarks = names
End Function

' The STYLEREF heading prefixes of the numbered equations from fromPosition on
' (automatic chapter numbers; written as placeholders by the shared writer).
Public Sub VTV2UpdateStyleRefsFrom(ByVal doc As Document, ByVal fromPosition As Long)
    Dim candidate As Field

    For Each candidate In doc.Range(fromPosition, doc.Content.End).Fields
        If candidate.Type = wdFieldStyleRef Then
            If Not VTFieldNativeMathRoot(candidate) Is Nothing Then candidate.Update
        End If
    Next candidate
End Sub
