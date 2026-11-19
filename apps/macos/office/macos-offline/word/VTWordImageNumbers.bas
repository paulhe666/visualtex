Attribute VB_Name = "VTWordImageNumbers"
Option Explicit
Option Private Module

' Numbers of image formulas: their SEQ fields, PlaceRef prefixes, static number layout and edits.

Private Const VT_WORD_NUMBERED_IMAGE_STYLE_NAME As String = _
    "VisualTeX Numbered Equation"

' The Word range a format conversion replaces for one image formula. A
' numbered image owns its paragraph: a tab, the image, a tab and the
' MACROBUTTON number. Replacing only the image left that number beside the new
' formula, so the number and the spacing around it go too. Text, fields or
' formulas elsewhere in the paragraph stay; paragraphCleared says whether the
' whole paragraph content was taken.
Public Function VTNumberedImageReplacementRange( _
    ByVal sourceRange As Range, _
    ByRef paragraphCleared As Boolean) As Range

    Dim documentObject As Document
    Dim placeRef As Field
    Dim paragraphRange As Range
    Dim replacementRange As Range
    Dim contentEnd As Long
    Dim numberStart As Long
    Dim numberEnd As Long

    paragraphCleared = False
    Set VTNumberedImageReplacementRange = sourceRange
    If sourceRange.InlineShapes.Count <> 1 Then Exit Function
    If VTRangeIsWithinTable(sourceRange) Then Exit Function
    Set placeRef = VTImagePlaceRef(sourceRange.InlineShapes(1).Range)
    If placeRef Is Nothing Then Exit Function
    Set documentObject = sourceRange.Document
    Set paragraphRange = sourceRange.Paragraphs(1).Range.Duplicate
    contentEnd = paragraphRange.End
    If documentObject.Range(contentEnd - 1, contentEnd).Text = vbCr Then
        contentEnd = contentEnd - 1
    End If
    numberStart = VTEquationFieldStart(placeRef)
    numberEnd = VTEquationFieldEnd(placeRef)
    If numberStart < sourceRange.End Or numberEnd > contentEnd Then Exit Function
    If Not VTRangeIsBlankSpacing( _
       documentObject.Range(sourceRange.End, numberStart)) Then Exit Function

    Set replacementRange = documentObject.Range(sourceRange.Start, numberEnd)
    If VTRangeIsBlankSpacing( _
       documentObject.Range(paragraphRange.Start, sourceRange.Start)) Then
        replacementRange.Start = paragraphRange.Start
    End If
    If VTRangeIsBlankSpacing( _
       documentObject.Range(numberEnd, contentEnd)) Then
        replacementRange.End = contentEnd
    End If
    paragraphCleared = (replacementRange.Start = paragraphRange.Start And _
        replacementRange.End = contentEnd)
    Set VTNumberedImageReplacementRange = replacementRange
End Function

Public Function VTMoveDirectImageEquationSequenceToExternalHelper( _
    ByRef formulaShape As InlineShape, _
    ByVal formulaId As String) As Field

    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim formulaContainerRange As Range
    Dim suffixRange As Range
    Dim helperParagraph As Range
    Dim fieldRange As Range
    Dim formattedSequenceRange As Range
    Dim directField As Field
    Dim helperField As Field
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim captionBookmarkName As String
    Dim preservedFieldCode As String
    Dim formulaStart As Long
    Dim helperStart As Long
    Dim imageMetadata As String
    Dim imageFormulaReference As String

    If formulaShape Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The image-to-native sequence migration target is invalid."
    End If
    Set documentObject = formulaShape.Range.Document
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    imageMetadata = formulaShape.AlternativeText
    imageFormulaReference = formulaShape.Title
    formulaStart = formulaShape.Range.Start
    Set directField = VTFindEquationSequenceField(paragraphRange)
    If directField Is Nothing Then
        Set helperField = VTNativeEquationSequenceHelperField( _
            documentObject, formulaId)
        If Not helperField Is Nothing Then
            Set VTMoveDirectImageEquationSequenceToExternalHelper = helperField
            Exit Function
        End If
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The numbered image has no direct Equation SEQ to migrate."
    End If

    preservedFieldCode = Trim$(directField.Code.Text)
    sequenceBookmarkName = VTEquationSequenceNumberBookmarkName(formulaId)
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)

    If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        documentObject.Bookmarks(sequenceBookmarkName).Delete
    End If
    If documentObject.Bookmarks.Exists(captionBookmarkName) Then
        documentObject.Bookmarks(captionBookmarkName).Delete
    End If

    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set suffixRange = documentObject.Range( _
        Start:=formulaContainerRange.End, End:=paragraphRange.End - 1)
    If suffixRange.End > suffixRange.Start Then suffixRange.Delete

    Set formulaShape = VTResolveImageFormulaByIdentity( _
        documentObject, formulaId, imageMetadata, _
        imageFormulaReference, formulaStart)
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    helperStart = paragraphRange.End
    Set helperParagraph = VTInsertEquationHelperParagraphAfter(paragraphRange)
    Set fieldRange = documentObject.Range( _
        Start:=helperStart, End:=helperStart)
    Set helperField = fieldRange.Fields.Add( _
        Range:=fieldRange, Type:=wdFieldEmpty, _
        Text:=preservedFieldCode, PreserveFormatting:=False)
    If VTFirstPositiveIntegerInText( _
       VTEquationSequenceResultText(helperField)) < 1 Then
        helperField.Update
    End If
    If Not VTIsNativeEquationSequenceField( _
       helperField, VTNativeEquationLabelName()) Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "Word did not migrate the image SEQ into an external helper."
    End If

    Set helperParagraph = _
        helperField.Result.Paragraphs(1).Range.Duplicate
    Set formattedSequenceRange = VTRefreshFormattedSequenceBookmark( _
        documentObject, helperField, formulaId)
    Set helperField = VTEquationSequenceFieldForBookmark( _
        documentObject, sequenceBookmarkName)
    If helperField Is Nothing Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The migrated image helper lost its VT_N_ identity."
    End If
    Set helperParagraph = _
        helperField.Result.Paragraphs(1).Range.Duplicate
    VTSetCollapsedEquationCaptionBookmark _
        documentObject, formulaId, helperParagraph
    VTFormatHiddenEquationParagraph helperParagraph
    If documentObject.Bookmarks.Exists(numberBookmarkName) Then
        documentObject.Bookmarks(numberBookmarkName).Delete
    End If

    Set VTMoveDirectImageEquationSequenceToExternalHelper = helperField
End Function

Public Function VTImageEquationNumberRange( _
    ByVal formulaRange As Range, _
    ByVal numberField As Field) As Range

    Dim exactRange As Range
    Dim containerRange As Range
    Dim paragraphRange As Range
    Dim fieldStart As Long
    Dim fieldEnd As Long

    If VTIsImagePlaceRef(numberField) Then
        Set VTImageEquationNumberRange = formulaRange.Document.Range( _
            VTEquationFieldStart(numberField), VTEquationFieldEnd(numberField))
        Exit Function
    End If
    If formulaRange Is Nothing Or numberField Is Nothing Then Exit Function
    If formulaRange.InlineShapes.Count <> 1 Or _
       formulaRange.OMaths.Count <> 0 Or _
       numberField.Type <> wdFieldRef Then Exit Function
    Set exactRange = formulaRange.InlineShapes(1).Range.Duplicate
    Set containerRange = _
        VTVisualTeXImageContainerRange(formulaRange.InlineShapes(1))
    Set paragraphRange = exactRange.Paragraphs(1).Range.Duplicate
    fieldStart = VTEquationFieldStart(numberField)
    fieldEnd = VTEquationFieldEnd(numberField)
    If fieldStart <= containerRange.End Or _
       fieldStart < 2 Or fieldEnd >= paragraphRange.End Or _
       exactRange.Document.Range(fieldStart - 2, fieldStart - 1).Text <> vbTab Or _
       exactRange.Document.Range(fieldStart - 1, fieldStart).Text <> "(" Or _
       exactRange.Document.Range(fieldEnd, fieldEnd + 1).Text <> ")" Then
        Exit Function
    End If
    Set VTImageEquationNumberRange = exactRange.Document.Range( _
        Start:=fieldStart - 1, End:=fieldEnd + 1)
End Function

Public Function VTStaticImageEquationNumberRange( _
    ByVal formulaRange As Range, _
    ByVal formulaId As String) As Range

    Dim documentObject As Document
    Dim exactRange As Range
    Dim containerRange As Range
    Dim paragraphRange As Range
    Dim numberRange As Range
    Dim separatorRange As Range
    Dim visibleNumberField As Field
    Dim bookmarkName As String
    Dim visibleText As String

    If formulaRange Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then Exit Function
    If formulaRange.InlineShapes.Count <> 1 Or _
       formulaRange.OMaths.Count <> 0 Then Exit Function

    Set documentObject = formulaRange.Document

    ' Current numbered images use an external SEQ helper plus one visible REF in
    ' the center/right-tab formula row. Keep this historical function name only
    ' as a compatibility resolver for regressions and older migration callers.
    Set visibleNumberField = VTImageEquationReferenceField( _
        formulaRange, formulaId)
    If Not visibleNumberField Is Nothing Then
        Set numberRange = VTImageEquationNumberRange( _
            formulaRange, visibleNumberField)
        If Not numberRange Is Nothing Then
            Set VTStaticImageEquationNumberRange = numberRange.Duplicate
            Exit Function
        End If
    End If

    bookmarkName = VTEquationNumberBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(bookmarkName) Then Exit Function
    Set exactRange = formulaRange.InlineShapes(1).Range.Duplicate
    Set containerRange = _
        VTVisualTeXImageContainerRange(formulaRange.InlineShapes(1))
    Set paragraphRange = exactRange.Paragraphs(1).Range.Duplicate
    Set numberRange = documentObject.Bookmarks( _
        bookmarkName).Range.Duplicate
    If numberRange.Information(wdWithInTable) Or _
       numberRange.Start <= containerRange.End Or _
       numberRange.End >= paragraphRange.End Or _
       numberRange.Fields.Count <> 1 Then Exit Function
    Set separatorRange = documentObject.Range( _
        Start:=containerRange.End, End:=numberRange.Start)
    If separatorRange.Text <> vbTab Then Exit Function
    visibleText = numberRange.Text
    If Len(visibleText) < 3 Or _
       Left$(visibleText, 1) <> "(" Or _
       Right$(visibleText, 1) <> ")" Or _
       VTLastIntegerInText(visibleText) < 1 Or _
       Not VTIsNativeEquationSequenceField( _
           numberRange.Fields(1), VTNativeEquationLabelName()) Then Exit Function
    Set VTStaticImageEquationNumberRange = numberRange.Duplicate
End Function

Public Sub VTApplyStaticImageEquationNumberFormatting( _
    ByVal numberRange As Range, _
    ByVal numberSize As Single, _
    Optional ByVal fontName As String = "")

    If numberRange Is Nothing Then
        Err.Raise vbObjectError + 7564, "VisualTeX", _
            "The static image Equation number formatting target is missing."
    End If
    If numberSize <= 0! Or _
       numberSize > CSng(VT_WORD_MAX_FORMULA_FONT_SIZE_PT) Then
        numberSize = VTVisibleEquationNumberFontSize(numberRange.Document)
    End If
    ' A genuine center/right-tab equation number is ordinary Word text. Match
    ' MathType and the Windows VisualTeX OLE path: keep the number on the normal
    ' paragraph baseline and move only the formula object for vertical alignment.
    With numberRange.Font
        If .Hidden <> False Then .Hidden = False
        If .Superscript <> False Then .Superscript = False
        If .Subscript <> False Then .Subscript = False
        If .Position <> 0 Then .Position = 0
    End With
    ' A number the user formatted keeps its size, font and colour while the
    ' caret moves or the document opens.
    If VT_WORD_KEEP_NUMBER_FORMAT Then Exit Sub
    If Len(fontName) = 0 Then fontName = VT_WORD_EQUATION_NUMBER_FONT_NAME
    With numberRange.Font
        If .Color <> wdColorAutomatic Then .Color = wdColorAutomatic
        If .Name <> fontName Then .Name = fontName
        If .NameAscii <> fontName Then .NameAscii = fontName
        If .NameOther <> fontName Then .NameOther = fontName
        If Abs(.Size - numberSize) > 0.1 Then .Size = numberSize
    End With
End Sub

' Word for Mac biases text the object model inserts towards East Asian
' (w:hint="eastAsia") when the text holds the U+2010 separator and a Chinese
' input source is selected. Word then draws the separator in the East Asian
' font, full width when that is SimSun (every document WPS makes): "(0 - 1)".
' NameFarEast cannot help, Word for Mac refuses a Latin East Asian font (error
' 5844). InsertSymbol with the default font bias writes the same character as
' ordinary text, so the number writers put every separator they wrote back
' that way, in the font of the character before it. Only where Word edits
' in place: in plain text, or inside a MACROBUTTON whose code is shown (with
' the code hidden Word inserts the symbol at the field start instead). With
' Track Changes on the replaced character would stay as a deletion.
Public Sub VTUnbiasNumberSeparators( _
    ByVal documentObject As Document, _
    ByVal startPosition As Long, _
    ByVal endPosition As Long)

    Dim position As Long
    Dim character As Range
    Dim fontName As String
    Dim contentEnd As Long

    If documentObject.TrackRevisions Then Exit Sub
    If startPosition < 1 Then startPosition = 1
    For position = startPosition To endPosition - 1
        Set character = documentObject.Range(position, position + 1)
        If character.Text = ChrW(8208) Then
            fontName = documentObject.Range(position - 1, position).Font.Name
            contentEnd = documentObject.Content.End
            If Len(fontName) > 0 And Left$(fontName, 1) <> "+" Then
                character.InsertSymbol CharacterNumber:=8208, Font:=fontName, _
                    Unicode:=True, Bias:=wdFontBiasDefault
            Else
                character.InsertSymbol CharacterNumber:=8208, _
                    Unicode:=True, Bias:=wdFontBiasDefault
            End If
            If documentObject.Content.End <> contentEnd Or _
               documentObject.Range(position, position + 1).Text <> ChrW(8208) Then
                Err.Raise vbObjectError + 7566, "VisualTeX", _
                    "Word did not rewrite the Equation number separator in place."
            End If
        End If
    Next position
End Sub

' True when a number prefix ends in a separator Word draws in the East Asian
' font (see VTUnbiasNumberSeparators): Font.Name is the font Word draws a
' character in, here neither the separator's own Latin font nor the digit's.
Private Function VTNumberSeparatorIsBiased(ByVal prefixRange As Range) As Boolean
    Dim documentObject As Document
    Dim separator As Range
    Dim drawnFont As String

    If prefixRange.End - prefixRange.Start < 2 Then Exit Function
    Set documentObject = prefixRange.Document
    If documentObject.TrackRevisions Then Exit Function
    Set separator = documentObject.Range(prefixRange.End - 1, prefixRange.End)
    If separator.Text <> ChrW(8208) Then Exit Function
    drawnFont = separator.Font.Name
    VTNumberSeparatorIsBiased = _
        drawnFont <> documentObject.Range( _
            prefixRange.End - 2, prefixRange.End - 1).Font.Name And _
        drawnFont <> separator.Font.NameOther
End Function

Public Function VTNumberedImageInkCenterYRatio( _
    ByVal formulaShape As InlineShape) As Double

    Dim documentObject As Document
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim inkCenterYRatio As Double

    inkCenterYRatio = 0.5
    If formulaShape Is Nothing Then
        VTNumberedImageInkCenterYRatio = inkCenterYRatio
        Exit Function
    End If
    Set documentObject = formulaShape.Range.Document
    If VTTryParseFormulaReference( _
       formulaShape.Title, formulaId, displayMode, numbered) Then
        VTTryReadWordImageInkCenterYRatio _
            documentObject, formulaId, inkCenterYRatio
    End If
    If inkCenterYRatio < 0# Or inkCenterYRatio > 1# Then _
        inkCenterYRatio = 0.5
    VTNumberedImageInkCenterYRatio = inkCenterYRatio
End Function

Private Function VTExpectedNumberedImageFormulaPosition( _
    ByVal formulaShape As InlineShape, _
    ByVal numberRange As Range, _
    ByVal valueLabel As String) As Long

    Dim numberSize As Double

    If formulaShape Is Nothing Or numberRange Is Nothing Then
        Err.Raise vbObjectError + 7564, "VisualTeX", _
            valueLabel & " target is missing."
    End If
    numberSize = CDbl(numberRange.Font.Size)
    If numberSize <= 0# Or _
       numberSize > VT_WORD_MAX_FORMULA_FONT_SIZE_PT Then
        numberSize = CDbl(VTVisibleEquationNumberFontSize( _
            formulaShape.Range.Document))
    End If
    ' Centre on the number as it is drawn: its own size and font.
    VTExpectedNumberedImageFormulaPosition = _
        VTExpectedImageFormulaPosition( _
            formulaShape, numberSize, valueLabel, CStr(numberRange.Font.Name))
End Function

Public Function VTApplyNumberedImageFormulaVerticalAlignment( _
    ByVal formulaShape As InlineShape, _
    ByVal numberRange As Range, _
    ByVal numberSize As Single, _
    ByVal valueLabel As String, _
    Optional ByVal fontName As String = "") As Long

    Dim formulaPosition As Long

    If formulaShape Is Nothing Or numberRange Is Nothing Then
        Err.Raise vbObjectError + 7564, "VisualTeX", _
            "The numbered image Equation vertical-alignment target is missing."
    End If

    VTNormalizeNumberedImageTextScaffold formulaShape, numberRange, numberSize, fontName
    formulaPosition = VTExpectedNumberedImageFormulaPosition( _
        formulaShape, numberRange, valueLabel)
    VTApplyWordImagePosition formulaShape, formulaPosition

    VTApplyNumberedImageFormulaVerticalAlignment = Abs(formulaPosition)
End Function

Private Sub VTNormalizeNumberedImageTextScaffold( _
    ByVal formulaShape As InlineShape, _
    ByVal numberRange As Range, _
    ByVal numberSize As Single, _
    Optional ByVal fontName As String = "")

    Dim paragraphRange As Range
    Dim characterRange As Range
    Dim oldStyleName As String
    Dim equationStyle As Style

    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    oldStyleName = VTParagraphStyleName(paragraphRange)
    If oldStyleName = "VisualTeX Numbered Equation" Or _
       oldStyleName = VT_WORD_NUMBERED_IMAGE_STYLE_NAME Then
        ' Historical numbered styles lowered every inherited character by the
        ' picture descent. Tabs kept that lowering even after the picture and
        ' number acquired independent positions. Retire that style contract.
        ' Repair the existing style in place. Reassigning Range.Style resets
        ' the paragraph's direct center/right tab stops on Word for Mac.
        Set equationStyle = paragraphRange.Style
        ' Rewriting BaseStyle/NextParagraphStyle also makes Mac Word drop the
        ' paragraph's direct tab overrides. Only the inherited lowering is old.
        With equationStyle.Font
            If .Superscript <> False Then .Superscript = False
            If .Subscript <> False Then .Subscript = False
            If .Position <> 0 Then .Position = 0
        End With
    End If
    VTApplyStaticImageEquationNumberFormatting numberRange, numberSize, fontName
    For Each characterRange In paragraphRange.Characters
        Select Case characterRange.Text
            Case vbTab, vbCr, Chr$(11)
                VTApplyStaticImageEquationNumberFormatting _
                    characterRange, numberSize, fontName
        End Select
    Next characterRange
End Sub

Public Function VTCalculateStaticImageEquationNumberPosition( _
    ByVal formulaShape As InlineShape, _
    ByVal numberRange As Range) As Long

    Dim documentObject As Document
    Dim numberSize As Single
    Dim numberPosition As Long

    If formulaShape Is Nothing Or numberRange Is Nothing Then
        Err.Raise vbObjectError + 7564, "VisualTeX", _
            "The static image Equation alignment target is missing."
    End If

    Set documentObject = formulaShape.Range.Document
    numberSize = numberRange.Font.Size
    If numberSize <= 0! Or _
       numberSize > CSng(VT_WORD_MAX_FORMULA_FONT_SIZE_PT) Then
        numberSize = VTVisibleEquationNumberFontSize(documentObject)
    End If
    numberPosition = VTApplyNumberedImageFormulaVerticalAlignment( _
        formulaShape, numberRange, numberSize, _
        "The image Equation visual-center position")
    ' Return only the positive magnitude for compatibility diagnostics. The
    ' shared routine applies the signed Position to the picture and keeps the
    ' complete visible number Range at Position 0.
    VTCalculateStaticImageEquationNumberPosition = numberPosition
End Function

Public Sub VTRefreshNumberedImageFormulaFontLayout( _
    ByRef formulaShape As InlineShape, _
    ByVal formulaId As String, _
    ByVal requestedFontSizePt As Double)

    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim numberRange As Range
    Dim numberField As Field
    Dim formulaStart As Long
    Dim imageMetadata As String
    Dim imageFormulaReference As String
    Dim expectedPosition As Long
    Dim numberFontSizePt As Single
    Dim numberFontName As String

    If formulaShape Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Or _
       Not VTValidWordFormulaFontSize(requestedFontSizePt) Then
        Err.Raise vbObjectError + 7571, "VisualTeX", _
            "The numbered image formula resize target is invalid."
    End If
    Set documentObject = formulaShape.Range.Document
    formulaStart = formulaShape.Range.Start
    imageMetadata = formulaShape.AlternativeText
    imageFormulaReference = formulaShape.Title
    ' The number follows the formula: its size and the font of its letters.
    numberFontSizePt = CSng(requestedFontSizePt)
    If numberFontSizePt > CSng(VT_WORD_MAX_FORMULA_FONT_SIZE_PT) Then _
        numberFontSizePt = CSng(VT_WORD_MAX_FORMULA_FONT_SIZE_PT)
    VTEnsureWordImageLetterFont formulaShape
    numberFontName = VTEquationNumberFontNameForFormula(documentObject, formulaId)
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate

    If paragraphRange.Information(wdWithInTable) Then
        ' A direct Word resize can reach an old document before the editor-open
        ' migration runs. Flatten only this selected legacy table now; after this
        ' call the normal resize path is identical to every new image formula.
        Set paragraphRange = VTMigrateLegacyNumberedImageTableToParagraph( _
            formulaShape, formulaId, "VisualTeX formula")
        Set formulaShape = VTResolveImageFormulaByIdentity( _
            documentObject, formulaId, imageMetadata, _
            imageFormulaReference, formulaStart)
        formulaStart = formulaShape.Range.Start
        Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    End If

    ' Rebuild both tab stops on every resize. The first keeps the image at the
    ' exact text-column centre; the second keeps the number at the right edge.
    VTConfigureNumberedEquationParagraph paragraphRange
    Set formulaShape = VTResolveImageFormulaByIdentity( _
        documentObject, formulaId, imageMetadata, _
        imageFormulaReference, formulaStart)
    Set numberField = VTImageEquationReferenceField( _
        formulaShape.Range.Duplicate, formulaId)
    If numberField Is Nothing Then
        Err.Raise vbObjectError + 7571, "VisualTeX", _
            "The numbered image formula has no visible number REF."
    End If
    If VTIsImagePlaceRef(numberField) Then
        VTUpdateImagePlaceRef numberField
    Else
        numberField.Update
    End If
    Set numberRange = VTImageEquationNumberRange( _
        formulaShape.Range.Duplicate, numberField)
    If numberRange Is Nothing Then
        Err.Raise vbObjectError + 7571, "VisualTeX", _
            "The numbered image formula has no complete visible REF range."
    End If
    VTFormatVisibleEquationReference _
        documentObject, numberField, numberRange, numberFontName, numberFontSizePt

    ' Keep the visible number at Word's ordinary paragraph baseline. Move only
    ' the formula object, using cached PNG alpha bounds and the number font's
    ' glyph bounds so their actual painted bbox centres coincide.
    expectedPosition = VTApplyNumberedImageFormulaVerticalAlignment( _
        formulaShape, numberRange, numberFontSizePt, _
        "The resized image Equation visual-center position", numberFontName)
    VTSetEquationNumberBookmarkExact _
        documentObject, formulaId, numberRange
    Set formulaShape = VTResolveImageFormulaByIdentity( _
        documentObject, formulaId, imageMetadata, _
        imageFormulaReference, formulaStart)
End Sub

Public Sub VTRefreshNumberedImageFormulaAfterGeometryChange( _
    ByRef formulaShape As InlineShape, _
    ByVal formulaId As String, _
    ByVal fontSizePt As Double)

    Dim refreshErrorNumber As Long
    Dim refreshErrorDescription As String

    On Error GoTo RefreshFailed
    VTBeginWordInternalMutation
    VTRefreshNumberedImageFormulaFontLayout _
        formulaShape, formulaId, fontSizePt
    VTEndWordInternalMutation
    Exit Sub

RefreshFailed:
    refreshErrorNumber = Err.Number
    refreshErrorDescription = Err.Description
    VTEndWordInternalMutation
    Err.Raise refreshErrorNumber, "VisualTeX image geometry", _
        refreshErrorDescription
End Sub

Public Function VTEnsureNumberedImageParagraphStyle( _
    ByVal documentObject As Document) As Style

    Dim equationStyle As Style

    If documentObject Is Nothing Then
        Err.Raise vbObjectError + 7564, "VisualTeX", _
            "The numbered image Equation style document is missing."
    End If
    If Not VT_WORD_DISPLAY_NUMBERING_IMAGE_STYLE Is Nothing Then
        Set VTEnsureNumberedImageParagraphStyle = VT_WORD_DISPLAY_NUMBERING_IMAGE_STYLE
        Exit Function
    End If
    On Error Resume Next
    Set equationStyle = documentObject.Styles( _
        VT_WORD_NUMBERED_IMAGE_STYLE_NAME)
    On Error GoTo StyleFailed
    If equationStyle Is Nothing Then
        Set equationStyle = documentObject.Styles.Add( _
            Name:=VT_WORD_NUMBERED_IMAGE_STYLE_NAME, _
            Type:=wdStyleTypeParagraph)
    End If
    equationStyle.BaseStyle = wdStyleNormal
    equationStyle.NextParagraphStyle = wdStyleNormal
    equationStyle.AutomaticallyUpdate = False
    With equationStyle.Font
        .Superscript = False
        .Subscript = False
        .Position = 0
    End With
    On Error Resume Next
    equationStyle.QuickStyle = False
    On Error GoTo StyleFailed
    Set VTEnsureNumberedImageParagraphStyle = equationStyle
    Exit Function

StyleFailed:
    Err.Raise vbObjectError + 7564, "VisualTeX", _
        "Word could not prepare the VisualTeX numbered Equation paragraph style."
End Function

Private Function VTParagraphStyleName(ByVal paragraphRange As Range) As String
    Dim attempt As Long
    On Error Resume Next
    For attempt = 1 To 3
        Err.Clear
        VTParagraphStyleName = paragraphRange.Style.NameLocal
        If Err.Number = 0 Then Exit Function
    Next attempt
    On Error GoTo 0
    VTParagraphStyleName = paragraphRange.Style.NameLocal
End Function

Public Function VTEnsureImageEquationNumber( _
    ByRef formulaShape As InlineShape, _
    ByVal renderedHeightPoints As Double, _
    ByVal formulaId As String, _
    ByVal captionText As String, _
    ByRef numberCreated As Boolean, _
    Optional ByVal deferReconcile As Boolean = False) As Range

    Dim formulaRange As Range
    Dim numberBookmarkName As String

    If formulaShape Is Nothing Then
        Err.Raise vbObjectError + 7502, "VisualTeX", _
            "The numbered formula image is missing."
    End If
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    numberCreated = Not formulaShape.Range.Document.Bookmarks.Exists( _
        numberBookmarkName)
    Set formulaRange = formulaShape.Range.Duplicate

    ' Tables are compatibility input only, never a retained image-number layout.
    ' Migrating the selected table is local; healthy legacy formulas keep their
    ' existing external VT_N_ helper and therefore stay on the O(1) edit path.
    If formulaRange.Information(wdWithInTable) Then
        Set VTEnsureImageEquationNumber = _
            VTMigrateLegacyNumberedImageTableToParagraph( _
                formulaShape, formulaId, captionText)
        Exit Function
    End If

    Set VTEnsureImageEquationNumber = VTInsertEquationNumber( _
        formulaShape, formulaId, captionText, deferReconcile)
    If VTEnsureImageEquationNumber.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 7569, "VisualTeX", _
            "A numbered image Equation was created inside a table."
    End If
End Function

Public Function VTNativeEquationOwnedImageNumberTail( _
    ByVal equationRange As Range, _
    ByVal formulaId As String) As Range

    Dim documentObject As Document
    Dim numberRange As Range
    Dim gapRange As Range
    Dim candidateField As Field
    Dim sequenceField As Field
    Dim numberBookmarkName As String
    Dim sequenceBookmarkName As String
    Dim gapText As String
    Dim isOwnedTail As Boolean

    Set documentObject = equationRange.Document
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    sequenceBookmarkName = VTEquationSequenceNumberBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(numberBookmarkName) Then Exit Function
    Set numberRange = documentObject.Bookmarks(numberBookmarkName).Range.Duplicate
    If numberRange.Information(wdWithInTable) Or _
       numberRange.OMaths.Count <> 0 Or numberRange.InlineShapes.Count <> 0 Or _
       numberRange.Start < equationRange.End Or _
       numberRange.Paragraphs(1).Range.Start <> _
           equationRange.Paragraphs(1).Range.Start Then Exit Function
    If Len(numberRange.Text) < 3 Or _
       Left$(numberRange.Text, 1) <> "(" Or _
       Right$(numberRange.Text, 1) <> ")" Or _
       VTLastIntegerInText(numberRange.Text) <= 0 Then Exit Function

    ' Field-free numbers are legacy migration input. A live REF must be fully
    ' owned by VT_R_ and resolve to this formula's durable external SEQ helper.
    isOwnedTail = (numberRange.Fields.Count = 0)
    If numberRange.Fields.Count = 1 Then
        Set candidateField = numberRange.Fields(1)
        If candidateField.Type = wdFieldRef And _
           StrComp(VTReferenceTargetBookmarkName(candidateField.Code.Text), _
               sequenceBookmarkName, vbTextCompare) = 0 Then
            Set sequenceField = VTNativeEquationSequenceHelperField( _
                documentObject, formulaId)
            If Not sequenceField Is Nothing Then
                ' Chapter/section helpers keep their prefix inside VT_N_ but
                ' outside the SEQ field result (for example VT_N_="0-1" while
                ' SEQ.Result="1"). A visible REF resolves the complete VT_N_,
                ' so compare it with that durable Bookmark rather than the bare
                ' SEQ result.
                isOwnedTail = _
                    VTEquationFieldStart(candidateField) >= numberRange.Start And _
                    VTEquationFieldEnd(candidateField) <= numberRange.End And _
                    numberRange.Text = "(" & candidateField.Result.Text & ")" And _
                    VTEquationSequenceResultText(candidateField) = _
                        VTEquationNumberTextForFormula( _
                            documentObject, formulaId)
            End If
        End If
    End If
    If Not isOwnedTail Then Exit Function
    Set gapRange = documentObject.Range( _
        Start:=equationRange.End, End:=numberRange.Start)
    gapText = gapRange.Text
    gapText = Replace$(gapText, vbTab, "")
    gapText = Replace$(gapText, " ", "")
    gapText = Replace$(gapText, ChrW(160), "")
    gapText = Replace$(gapText, ChrW(8203), "")
    gapText = Replace$(gapText, ChrW(8288), "")
    If Len(gapText) <> 0 Then Exit Function
    Set VTNativeEquationOwnedImageNumberTail = documentObject.Range( _
        Start:=equationRange.End, End:=numberRange.End)
End Function

' Size and font of an image formula's Equation number: the formula's
' VisualTeX font size and the Word font of its letters.
Public Sub VTImageFormulaNumberFormat( _
    ByVal formulaShape As InlineShape, _
    ByRef fontName As String, _
    ByRef fontSize As Single)

    Dim documentObject As Document
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim formulaSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim observedWordFontSizePt As Double

    Set documentObject = formulaShape.Range.Document
    fontName = VT_WORD_EQUATION_NUMBER_FONT_NAME
    fontSize = VTVisibleEquationNumberFontSize(documentObject)
    If Not VTTryParseFormulaReference( _
       formulaShape.Title, formulaId, displayMode, numbered) Then Exit Sub
    fontName = VTEquationNumberFontNameForFormula(documentObject, formulaId)
    If VTTryReadWordImageScaleState(documentObject, formulaId, formulaSizePt, _
       referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
       observedWordFontSizePt) Then
        If formulaSizePt <= VT_WORD_MAX_FORMULA_FONT_SIZE_PT Then _
            fontSize = CSng(formulaSizePt)
    End If
End Sub

' Current Word formats mirror the Windows native host contracts. Image numbers
' live wholly in their formula paragraph. OMML has only native math/Word fields;
' transaction bookmarks and old VisualTeX metadata are removed before commit.
Public Function VTIsImagePlaceRef(ByVal candidate As Field) As Boolean
    If candidate Is Nothing Then Exit Function
    If candidate.Type <> wdFieldMacroButton Then Exit Function
    VTIsImagePlaceRef = (InStr(1, candidate.Code.Text, _
        "MACROBUTTON VisualTeXPlaceRef ", vbTextCompare) > 0)
End Function

Public Function VTImagePlaceRef(ByVal formulaRange As Range) As Field
    Dim candidate As Field
    If formulaRange Is Nothing Then Exit Function
    For Each candidate In formulaRange.Paragraphs(1).Range.Fields
        If VTIsImagePlaceRef(candidate) Then
            Set VTImagePlaceRef = candidate
            Exit Function
        End If
    Next candidate
End Function

Public Function VTImageSequenceInstruction( _
    ByVal documentObject As Document, _
    Optional ByVal hiddenIncrement As Boolean = False, _
    Optional ByVal currentValue As Boolean = False, _
    Optional ByVal restartAtOne As Boolean = False) As String
    Dim instruction As String
    instruction = "SEQ " & VT_WORD_IMAGE_SEQUENCE_NAME
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
    VTImageSequenceInstruction = instruction
End Function

Public Function VTImageSequenceField( _
    ByVal placeRef As Field, ByVal hiddenIncrement As Boolean) As Field
    Dim candidate As Field
    Dim isHidden As Boolean
    If Not VTIsImagePlaceRef(placeRef) Then Exit Function
    For Each candidate In placeRef.Code.Fields
        If candidate.Type = wdFieldSequence And _
           StrComp(VTSequenceIdentifierFromFieldCode(candidate.Code.Text), _
               VT_WORD_IMAGE_SEQUENCE_NAME, vbTextCompare) = 0 Then
            isHidden = VTEquationSequenceFieldHasSwitch(candidate, "h")
            If isHidden = hiddenIncrement Then
                Set VTImageSequenceField = candidate
                Exit Function
            End If
        End If
    Next candidate
End Function

Private Function VTImagePlaceRefIsCanonical( _
    ByVal placeRef As Field, _
    ByVal expectedRestartAtOne As Boolean, _
    ByVal expectedPrefixText As String) As Boolean
    Dim candidate As Field
    Dim codeText As String
    Dim sequenceCount As Long
    Dim hiddenCount As Long
    Dim currentCount As Long
    Dim prefixRange As Range
    Dim numberField As Field
    If Not VTIsImagePlaceRef(placeRef) Then Exit Function
    codeText = placeRef.Code.Text
    If Len(codeText) = 0 Or Right$(codeText, 1) <> ")" Then Exit Function
    For Each candidate In placeRef.Code.Fields
        If candidate.Type = wdFieldSequence Then
            If StrComp(VTSequenceIdentifierFromFieldCode(candidate.Code.Text), _
               VT_WORD_IMAGE_SEQUENCE_NAME, vbTextCompare) <> 0 Then Exit Function
            sequenceCount = sequenceCount + 1
            If VTEquationSequenceFieldHasSwitch(candidate, "h") Then
                hiddenCount = hiddenCount + 1
                If Not VTEquationSequenceFieldIsCompatible( _
                   candidate, VT_WORD_IMAGE_SEQUENCE_NAME, True, False, _
                   expectedRestartAtOne) Then Exit Function
            ElseIf VTEquationSequenceFieldHasSwitch(candidate, "c") Then
                currentCount = currentCount + 1
                Set numberField = candidate
                If Not VTEquationSequenceFieldIsCompatible( _
                   candidate, VT_WORD_IMAGE_SEQUENCE_NAME, False, True, _
                   False) Then Exit Function
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    Next candidate
    If sequenceCount <> 2 Or hiddenCount <> 1 Or currentCount <> 1 Then Exit Function
    Set prefixRange = VTImageNumberPrefixRange(placeRef, numberField)
    If prefixRange Is Nothing Then Exit Function
    prefixRange.TextRetrievalMode.IncludeFieldCodes = False
    If StrComp(prefixRange.Text, expectedPrefixText, _
       vbBinaryCompare) <> 0 Then Exit Function
    ' Update Numbers and format changes rewrite a number whose separator Word
    ' draws in the East Asian font (written before VTUnbiasNumberSeparators).
    If VTNumberSeparatorIsBiased(prefixRange) Then Exit Function
    VTImagePlaceRefIsCanonical = True
End Function

Public Function VTCreateCanonicalImagePlaceRef( _
    ByVal documentObject As Document, _
    ByVal insertionPosition As Long, _
    ByVal restartAtOne As Boolean, _
    ByVal prefixText As String) As Field
    Dim insertion As Range
    Dim codeRange As Range
    Dim placeRef As Field
    Dim numberField As Field
    Dim increment As Field
    Dim nativeCodeText As String
    Dim completedCodeText As String

    Set insertion = documentObject.Range(insertionPosition, insertionPosition)
    Set placeRef = documentObject.Fields.Add( _
        Range:=insertion, Type:=wdFieldMacroButton, _
        Text:="VisualTeXPlaceRef", PreserveFormatting:=False)
    placeRef.ShowCodes = True

    ' Word creates the required separator after the MACROBUTTON instruction.
    ' Make the real terminal punctuation the first label character inserted at
    ' Code.End, then build every preceding segment in reverse at that fixed
    ' interior position. The native separator remains before the label and the
    ' completed field therefore ends at ')' with no synthetic trailing space.
    Set codeRange = placeRef.Code.Duplicate
    nativeCodeText = codeRange.Text
    If Len(nativeCodeText) = 0 Or _
       InStr(1, " " & vbTab & vbCr & vbLf, _
           Right$(nativeCodeText, 1), vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7605, "VisualTeX", _
            "Word did not create the native MACROBUTTON instruction separator."
    End If
    insertionPosition = codeRange.End
    Set insertion = documentObject.Range(insertionPosition, insertionPosition)
    insertion.InsertAfter ")"

    Set codeRange = placeRef.Code.Duplicate
    If insertionPosition <= codeRange.Start Or _
       insertionPosition >= codeRange.End Then
        Err.Raise vbObjectError + 7605, "VisualTeX", _
            "Word did not retain the image-number closing parenthesis inside MACROBUTTON."
    End If

    Set insertion = documentObject.Range(insertionPosition, insertionPosition)
    Set numberField = VTInsertFieldBefore(insertion, _
        VTImageSequenceInstruction(documentObject, False, True, False))
    If Len(prefixText) > 0 Then
        Set insertion = documentObject.Range(insertionPosition, insertionPosition)
        insertion.Text = prefixText
        VTUnbiasNumberSeparators documentObject, _
            insertionPosition, insertionPosition + Len(prefixText)
    End If
    Set insertion = documentObject.Range(insertionPosition, insertionPosition)
    insertion.Text = "("
    Set insertion = documentObject.Range(insertionPosition, insertionPosition)
    Set increment = VTInsertFieldBefore(insertion, _
        VTImageSequenceInstruction(documentObject, True, False, restartAtOne))

    On Error Resume Next
    increment.ShowCodes = False
    numberField.ShowCodes = False
    On Error GoTo 0
    completedCodeText = placeRef.Code.Text
    If Len(completedCodeText) = 0 Or Right$(completedCodeText, 1) <> ")" Then
        Err.Raise vbObjectError + 7605, "VisualTeX", _
            "Word left trailing whitespace after the image Equation number."
    End If
    On Error Resume Next
    placeRef.ShowCodes = False
    On Error GoTo 0
    Set VTCreateCanonicalImagePlaceRef = placeRef
End Function

Private Function VTImageNumberPrefixRange( _
    ByVal placeRef As Field, ByVal numberField As Field) As Range

    Dim documentObject As Document
    Dim codeStart As Long
    Dim fieldStart As Long
    Dim characterPosition As Long
    Dim characterRange As Range
    Dim prefixRange As Range

    If Not VTIsImagePlaceRef(placeRef) Or numberField Is Nothing Then _
        Exit Function
    Set documentObject = placeRef.Code.Document
    codeStart = placeRef.Code.Start
    fieldStart = VTEquationFieldStart(numberField)
    If fieldStart <= codeStart Then Exit Function

    ' Range.Find is view-sensitive inside a MACROBUTTON code range on Word for
    ' Mac: after ShowCodes is restored it can skip literal label characters
    ' around nested fields. The label is short, so resolve the actual document
    ' character boundary directly and take the last opening parenthesis before
    ' the current-value SEQ field.
    For characterPosition = fieldStart - 1 To codeStart Step -1
        Set characterRange = documentObject.Range( _
            characterPosition, characterPosition + 1)
        If characterRange.Text = "(" Then
            Set prefixRange = documentObject.Range( _
                characterPosition + 1, fieldStart)
            Set VTImageNumberPrefixRange = prefixRange
            Exit Function
        End If
    Next characterPosition
End Function

' The literal chapter prefix of a picture's number as it is written ("2.1-"
' in "(2.1-3)", nothing in "(3)"): what the document's numbering style is
' read from.
Public Function VTTryImageNumberPrefixText( _
    ByVal placeRef As Field, ByRef prefixText As String) As Boolean
    Dim numberField As Field
    Dim prefixRange As Range
    If Not VTIsImagePlaceRef(placeRef) Then Exit Function
    Set numberField = VTImageSequenceField(placeRef, False)
    If numberField Is Nothing Then Exit Function
    Set prefixRange = VTImageNumberPrefixRange(placeRef, numberField)
    If prefixRange Is Nothing Then Exit Function
    prefixRange.TextRetrievalMode.IncludeFieldCodes = False
    prefixText = prefixRange.Text
    VTTryImageNumberPrefixText = True
End Function

Public Function VTCurrentImageNumberRange(ByVal placeRef As Field) As Range
    Dim codeRange As Range
    Dim prefixRange As Range
    Dim numberField As Field
    If Not VTIsImagePlaceRef(placeRef) Then Exit Function
    Set numberField = VTImageSequenceField(placeRef, False)
    If numberField Is Nothing Then Exit Function
    Set prefixRange = VTImageNumberPrefixRange(placeRef, numberField)
    If prefixRange Is Nothing Then Exit Function
    Set codeRange = placeRef.Code.Duplicate
    codeRange.SetRange Start:=prefixRange.Start, End:=numberField.Result.End
    codeRange.TextRetrievalMode.IncludeFieldCodes = False
    Set VTCurrentImageNumberRange = codeRange
End Function

Public Sub VTUpdateImagePlaceRef(ByVal placeRef As Field)
    ' Never update the outer MACROBUTTON: its display lives in its instruction,
    ' exactly as in Word/MathType. Update nested fields in physical order,
    ' all in one call: a Field.Update costs 6 ms in Word for Mac, a Fields.Update
    ' of a collection 0.3 ms a field.
    Dim nested As Fields
    Set nested = placeRef.Code.Fields
    If nested.Count > 0 Then nested.Update
End Sub

Public Sub VTRetireExternalSequenceHelper( _
    ByVal documentObject As Document, ByVal formulaId As String, _
    ByVal numberSource As Range)
    Dim oldSequence As Field
    Dim helper As Range
    Dim candidate As Bookmark
    Dim targets As New Collection
    Dim item As Variant
    Dim name As String
    Dim showHidden As Boolean
    Set oldSequence = VTEquationSequenceFieldForBookmark( _
        documentObject, VTEquationSequenceNumberBookmarkName(formulaId))
    If Not oldSequence Is Nothing Then
        Set helper = oldSequence.Result.Paragraphs(1).Range.Duplicate
        If Not VTHelperParagraphOwnsNativeEquationSequence(helper) Then
            Set helper = Nothing
        End If
    End If
    If Not helper Is Nothing Then
        ' Include Word's _Ref bookmarks: existing native cross-references must
        ' follow their actual number when the former source paragraph disappears.
        showHidden = documentObject.Bookmarks.ShowHidden
        documentObject.Bookmarks.ShowHidden = True
        For Each candidate In helper.Bookmarks
            If candidate.Name <> VTEquationCaptionBookmarkName(formulaId) Then
                targets.Add candidate.Name
            End If
        Next candidate
        documentObject.Bookmarks.ShowHidden = showHidden
        For Each item In targets
            documentObject.Bookmarks.Add Name:=CStr(item), Range:=numberSource
        Next item
        helper.Delete
    End If
    name = VTEquationCaptionBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(name) Then documentObject.Bookmarks(name).Delete
End Sub

Public Function VTWriteSingleParagraphImageNumber( _
    ByRef formulaShape As InlineShape, _
    ByVal formulaId As String, _
    Optional ByVal headingPrefixOverride As String = "", _
    Optional ByVal headingPrefixResolved As Boolean = False, _
    Optional ByVal restartAtOneOverride As Boolean = False, _
    Optional ByVal restartResolved As Boolean = False, _
    Optional ByVal placeRefTemplate As Range = Nothing, _
    Optional ByVal updateNumberFields As Boolean = True) As Range
    Dim doc As Document
    Dim paragraphRange As Range
    Dim container As Range
    Dim insertion As Range
    Dim suffix As Range
    Dim placeRef As Field
    Dim numberSource As Range
    Dim fullNumber As Range
    Dim oldFullNumber As Range
    Dim preservedBookmarks As Collection
    Dim preservedName As Variant
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim level As Long
    Dim headingPrefix As String
    Dim prefixText As String
    Dim restartAtOne As Boolean
    Dim position As Long
    Dim numberFontName As String
    Dim numberFontSize As Single
    Dim discardedText As String
    Dim candidate As Field
    Set doc = formulaShape.Range.Document
    level = VTNativeHeadingLevelAtRange(formulaShape.Range)
    If level > 0 Then
        If headingPrefixResolved Then
            headingPrefix = headingPrefixOverride
        Else
            headingPrefix = VTEquationHeadingPrefix( _
                formulaShape.Range, VTEquationNumberingMode(doc))
        End If
        prefixText = headingPrefix & _
            VTEquationNumberingDisplaySeparator(doc)
        If restartResolved Then
            restartAtOne = restartAtOneOverride
        Else
            restartAtOne = VTEquationStartsNumberScope( _
                formulaShape.Range, True)
        End If
    End If
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    If VTRangeIsWithinTable(paragraphRange) Then
        Err.Raise vbObjectError + 7605, "VisualTeX", _
            "Migrate the legacy image table before attaching paragraph numbering."
    End If
    Set container = VTVisualTeXImageContainerRange(formulaShape)
    Set insertion = doc.Range(paragraphRange.Start, container.Start)
    If VTWordRangeHasMeaningfulText(insertion) Then
        Err.Raise vbObjectError + 7605, "VisualTeX", _
            "A numbered image requires a dedicated paragraph."
    End If
    If paragraphRange.OMaths.Count > 0 Then VTClearImageParagraphMathMark formulaShape
    insertion.Text = vbTab
    Set container = VTVisualTeXImageContainerRange(formulaShape)
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    VTConfigureNumberedEquationParagraph paragraphRange
    Set placeRef = VTImagePlaceRef(formulaShape.Range)
    If Not placeRef Is Nothing Then
        If Not VTImagePlaceRefIsCanonical( _
           placeRef, restartAtOne, prefixText) Then
            Set suffix = doc.Range(container.End, paragraphRange.End - 1)
            Set oldFullNumber = doc.Range( _
                VTEquationFieldStart(placeRef), VTEquationFieldEnd(placeRef))
            If oldFullNumber.Start < suffix.Start Or _
               oldFullNumber.End > suffix.End Then
                Err.Raise vbObjectError + 7605, "VisualTeX", _
                    "The legacy image number crosses its formula paragraph."
            End If
            For Each candidate In suffix.Fields
                If VTEquationFieldStart(candidate) < oldFullNumber.Start Or _
                   VTEquationFieldEnd(candidate) > oldFullNumber.End Then
                    Err.Raise vbObjectError + 7605, "VisualTeX", _
                        "The image number tail contains an unrelated field."
                End If
            Next candidate
            discardedText = doc.Range( _
                suffix.Start, oldFullNumber.Start).Text & doc.Range( _
                oldFullNumber.End, suffix.End).Text
            discardedText = Replace$(discardedText, vbTab, "")
            discardedText = Replace$(discardedText, " ", "")
            discardedText = Replace$(discardedText, ChrW(160), "")
            discardedText = Replace$(discardedText, ChrW(8203), "")
            discardedText = Replace$(discardedText, ChrW(8288), "")
            discardedText = Replace$(discardedText, Chr$(11), "")
            If Len(discardedText) > 0 Then
                Err.Raise vbObjectError + 7605, "VisualTeX", _
                    "The image number tail contains body text."
            End If
            Set preservedBookmarks = _
                VTCaptureImageNumberBookmarks(formulaShape, formulaId)
            oldFullNumber.Delete
            Set placeRef = Nothing
            Set container = VTVisualTeXImageContainerRange(formulaShape)
            Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
        End If
    End If
    If placeRef Is Nothing Then
        Set suffix = doc.Range(container.End, paragraphRange.End - 1)
        discardedText = suffix.Text
        For Each candidate In suffix.Fields
            If candidate.Type <> wdFieldRef And candidate.Type <> wdFieldSequence Then
                Err.Raise vbObjectError + 7605, "VisualTeX", _
                    "The image number tail contains an unrelated field."
            End If
            discardedText = Replace$(discardedText, candidate.Result.Text, "")
        Next candidate
        discardedText = Replace$(Replace$(discardedText, vbTab, ""), " ", "")
        discardedText = Replace$(Replace$(discardedText, "(", ""), ")", "")
        discardedText = Replace$(discardedText, ChrW(160), "")
        discardedText = Replace$(discardedText, ChrW(8203), "")
        discardedText = Replace$(discardedText, ChrW(8288), "")
        discardedText = Replace$(discardedText, Chr$(11), "")
        If Len(discardedText) > 0 And Not IsNumeric(discardedText) Then
            Err.Raise vbObjectError + 7605, "VisualTeX", _
                "The image number tail contains body text" & _
                " [raw=" & Replace$(Replace$(suffix.Text, vbCr, _
                    "<CR>"), vbTab, "<TAB>") & _
                "; fields=" & CStr(suffix.Fields.Count) & "]."
        End If
        suffix.Text = vbTab
        Set insertion = doc.Range(suffix.End, suffix.End)
        If placeRefTemplate Is Nothing Then
            Set placeRef = VTCreateCanonicalImagePlaceRef( _
                doc, insertion.Start, restartAtOne, prefixText)
        Else
            ' Batch numbering copies one prepared number (same prefix and
            ' restart) instead of adding its fields: Word evaluates every
            ' added SEQ field by counting all fields before it.
            insertion.FormattedText = placeRefTemplate.FormattedText
            Set placeRef = VTImagePlaceRef(formulaShape.Range)
            If placeRef Is Nothing Then Err.Raise vbObjectError + 7605, _
                "VisualTeX", "Word did not copy the image Equation number."
        End If
    End If
    Set numberSource = VTCurrentImageNumberRange(placeRef)
    If numberSource Is Nothing Then Err.Raise vbObjectError + 7605, _
        "VisualTeX", "The paragraph number has no visible SEQ source."
    VTRetireExternalSequenceHelper doc, formulaId, numberSource
    If updateNumberFields Then VTUpdateImagePlaceRef placeRef
    Set numberSource = VTCurrentImageNumberRange(placeRef)
    Set fullNumber = doc.Range(VTEquationFieldStart(placeRef), VTEquationFieldEnd(placeRef))
    sequenceBookmarkName = VTEquationSequenceNumberBookmarkName(formulaId)
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    If Not preservedBookmarks Is Nothing Then
        For Each preservedName In preservedBookmarks
            If StrComp(CStr(preservedName), sequenceBookmarkName, _
               vbTextCompare) <> 0 And _
               StrComp(CStr(preservedName), numberBookmarkName, _
               vbTextCompare) <> 0 Then
                If doc.Bookmarks.Exists(CStr(preservedName)) Then _
                    doc.Bookmarks(CStr(preservedName)).Delete
                doc.Bookmarks.Add Name:=CStr(preservedName), Range:=numberSource
            End If
        Next preservedName
    End If
    If doc.Bookmarks.Exists(sequenceBookmarkName) Then _
        doc.Bookmarks(sequenceBookmarkName).Delete
    doc.Bookmarks.Add Name:=sequenceBookmarkName, Range:=numberSource
    VTSetEquationNumberBookmarkExact doc, formulaId, fullNumber
    VTImageFormulaNumberFormat formulaShape, numberFontName, numberFontSize
    VTApplyStaticImageEquationNumberFormatting fullNumber, numberFontSize, numberFontName
    position = VTApplyNumberedImageFormulaVerticalAlignment( _
        formulaShape, fullNumber, numberFontSize, _
        "The same-paragraph image number baseline", numberFontName)
    Set VTWriteSingleParagraphImageNumber = fullNumber.Duplicate
End Function

Public Sub VTUpdateImageNumberFieldsFrom( _
    ByVal doc As Document, ByVal changedFrom As Long, _
    Optional ByVal updateReferences As Boolean = True)

    Dim scan As Range
    Dim paragraph As Paragraph
    Dim candidate As Field
    Dim startPosition As Long

    If doc Is Nothing Then Exit Sub
    startPosition = changedFrom
    If startPosition < doc.Content.Start Then startPosition = doc.Content.Start
    If startPosition > doc.Content.End Then startPosition = doc.Content.End
    Set scan = doc.Range(Start:=startPosition, End:=doc.Content.End)

    ' Re-evaluate only the affected suffix of the private image sequence. The
    ' outer MACROBUTTON result is instruction text, so update its nested hidden
    ' increment and current-value fields in physical document order.
    ' The suffix's own Fields collection lists the same fields in the same
    ' physical order as its paragraphs do, without one COM round trip per
    ' paragraph (about 0.25 ms each on Word for Mac).
    VTWordPerformanceMark "image-renumber-start"
    For Each candidate In scan.Fields
        If VTIsImagePlaceRef(candidate) And _
           VTEquationFieldStart(candidate) >= startPosition Then
            VTUpdateImagePlaceRef candidate
        End If
    Next candidate
    VTWordPerformanceMark "image-renumber-places-updated"

    ' References are cheap compared with rebuilding every formula and Word may
    ' cache their displayed results. Refresh them only after all affected image
    ' sources have reached their final values. This does not mutate either SEQ
    ' chain and therefore cannot couple image and OMML numbering again. A
    ' caller that renumbers both families refreshes them once, afterwards.
    If Not updateReferences Then Exit Sub
    For Each candidate In doc.Fields
        If candidate.Type = wdFieldRef Then candidate.Update
    Next candidate
    VTWordPerformanceMark "image-renumber-references-updated"
End Sub

Public Sub VTRemoveImageNumberingForEdit( _
    ByVal shape As InlineShape, ByVal formulaId As String)
    Dim doc As Document
    Dim placeRef As Field
    Dim ownedNumber As Range
    Set doc = shape.Range.Document
    Set placeRef = VTImagePlaceRef(shape.Range)
    If Not placeRef Is Nothing Then
        Set ownedNumber = doc.Range(VTEquationFieldStart(placeRef), VTEquationFieldEnd(placeRef))
        If ownedNumber.Start > shape.Range.End Then
            If doc.Range(ownedNumber.Start - 1, ownedNumber.Start).Text = vbTab Then _
                ownedNumber.Start = ownedNumber.Start - 1
        End If
        ownedNumber.Delete
    End If
    ' Copied formulas may have their private number field without its former
    ' UUID bookmarks. Remove the actual local field, then clean only this id's
    ' remaining helper identities. Never delete the original copy's scaffold.
    VTDeleteEquationNumberScaffold doc, formulaId, False, True
End Sub

Public Sub VTApplyImageNumberingEdit( _
    ByRef shape As InlineShape, ByVal formulaId As String, ByVal numbered As Boolean)
    Dim doc As Document
    Dim result As Range
    Dim prefix As Range
    Set doc = shape.Range.Document
    If numbered Then
        Set result = VTWriteSingleParagraphImageNumber(shape, formulaId)
    Else
        VTRemoveImageNumberingForEdit shape, formulaId
        Set prefix = shape.Range.Paragraphs(1).Range.Duplicate
        prefix.End = shape.Range.Start
        If prefix.Text = vbTab Then prefix.Delete
        VTNormalizeUnnumberedDisplayParagraph shape.Range
    End If
    VTNormalizeImageDisplayParagraph shape.Range
    VTReconcileNumberingAfterEdit doc, shape.Range.Start, True
End Sub

Public Function VTCaptureImageNumberBookmarks(ByVal shape As InlineShape, ByVal formulaId As String) As Collection
    Dim names As New Collection
    Dim number As Range
    Dim mark As Bookmark
    Dim doc As Document
    Dim hidden As Boolean
    Set doc = shape.Range.Document
    Set number = VTImageEquationNumberRange(shape.Range, VTImageEquationReferenceField(shape.Range, _
        formulaId))
    If number Is Nothing Then
        Set VTCaptureImageNumberBookmarks = names
        Exit Function
    End If
    hidden = doc.Bookmarks.ShowHidden
    doc.Bookmarks.ShowHidden = True
    For Each mark In doc.Bookmarks
        If mark.Range.Start >= number.Start And mark.Range.End <= number.End Then names.Add mark.Name
    Next mark
    doc.Bookmarks.ShowHidden = hidden
    Set VTCaptureImageNumberBookmarks = names
End Function

Public Sub VTRestoreImageNumberBookmarks( _
    ByVal formulaShape As InlineShape, ByVal names As Collection)

    Dim placeRef As Field
    Dim source As Range
    Dim item As Variant

    If formulaShape Is Nothing Or names Is Nothing Then Exit Sub
    Set placeRef = VTImagePlaceRef(formulaShape.Range)
    If placeRef Is Nothing Then Exit Sub
    Set source = VTCurrentImageNumberRange(placeRef)
    If source Is Nothing Then Exit Sub
    For Each item In names
        If formulaShape.Range.Document.Bookmarks.Exists(CStr(item)) Then _
            formulaShape.Range.Document.Bookmarks(CStr(item)).Delete
        formulaShape.Range.Document.Bookmarks.Add _
            Name:=CStr(item), Range:=source
    Next item
End Sub

' Sets the chapter restart (\r 1) of every numbered display image from
' fromPosition on by the rule of VTEquationStartsNumberScope: an image restarts
' when no numbered image precedes it or a heading at the restart level lies
' between the two. One walk over the document's images and the paragraphs
' between numbered images.
Public Sub VTRepairImageNumberRestartsFrom( _
    ByVal documentObject As Document, _
    ByVal fromPosition As Long)

    Dim shape As InlineShape
    Dim paragraph As Paragraph
    Dim placeRef As Field
    Dim hiddenField As Field
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim restartAtOne As Boolean
    Dim previousEnd As Long
    Dim targetLevel As Long
    Dim candidateLevel As Long

    If documentObject Is Nothing Then Exit Sub
    targetLevel = VTEquationNumberingRestartLevel(documentObject)
    If targetLevel < 1 Then Exit Sub

    previousEnd = -1
    For Each shape In documentObject.InlineShapes
        formulaId = ""
        displayMode = ""
        numbered = False
        If Not VTTryParseFormulaReference( _
           shape.Title, formulaId, displayMode, numbered) Then GoTo NextShape
        If Not numbered Or displayMode <> "block" Then GoTo NextShape
        If shape.Range.Start >= fromPosition Then
            If previousEnd < 0 Then
                restartAtOne = True
            ElseIf previousEnd >= shape.Range.Start Then
                restartAtOne = False
            Else
                restartAtOne = False
                For Each paragraph In documentObject.Range( _
                    previousEnd, shape.Range.Start).Paragraphs
                    candidateLevel = paragraph.OutlineLevel
                    If candidateLevel >= 1 And candidateLevel <= targetLevel Then
                        restartAtOne = True
                        Exit For
                    End If
                Next paragraph
            End If
            Set placeRef = VTImagePlaceRef(shape.Range)
            If Not placeRef Is Nothing Then
                Set hiddenField = VTImageSequenceField(placeRef, True)
                If hiddenField Is Nothing Then
                    Err.Raise vbObjectError + 7605, "VisualTeX", _
                        "A numbered image has no hidden sequence field."
                End If
                If VTEquationSequenceFieldHasSwitch(hiddenField, "r") <> _
                   restartAtOne Then
                    hiddenField.Code.Text = " " & VTImageSequenceInstruction( _
                        documentObject, True, False, restartAtOne) & " "
                End If
            End If
        End If
        previousEnd = shape.Range.Paragraphs(1).Range.End
NextShape:
    Next shape
End Sub
