Attribute VB_Name = "VTWordAdapter"
Option Explicit

Private Const VT_WORD_HOST As String = "word"
Private Const VT_WORD_STATUS_FILE As String = "/OfficePluginStatus/word.json"
Private Const VT_WORD_BOOKMARK_PREFIX As String = "VT_Pending_"
Private Const VT_WORD_NATIVE_BOOKMARK_PREFIX As String = "VT_F_"
Private Const VT_WORD_CAPTION_BOOKMARK_PREFIX As String = "VT_C_"
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

Public Sub AutoExec()
    On Error Resume Next
    VTInitializeWordEvents
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

Public Sub VisualTeX_RunWordNativeRegression()
    Const fixtureFormulaId As String = _
        "11111111-1111-4111-8111-111111111111"
    Const nativeFormulaId As String = _
        "22222222-2222-4222-8222-222222222222"
    Const conversionFormulaId As String = _
        "33333333-3333-4333-8333-333333333333"

    Dim testDocument As Document
    Dim placeholder As InlineShape
    Dim insertionRange As Range
    Dim equationRange As Range
    Dim numberRange As Range
    Dim crossReferenceItems As Variant
    Dim nativeDocumentPath As String
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
    Dim diagnosticRange As Range
    Dim crossReferenceTextFound As Boolean
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
    If Not VTPathFileExists(nativeDocumentPath) Then
        Err.Raise vbObjectError + 7482, "VisualTeX", _
            "The Word native regression DOCX fixture is missing."
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
    crossReferenceItems = _
        testDocument.GetCrossReferenceItems(wdCaptionEquation)
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
        If InStr(1, CStr(crossReferenceItems(itemIndex)), _
            "x^2 + y^2", vbTextCompare) > 0 Then
            crossReferenceTextFound = True
        End If
    Next itemIndex
    If Not crossReferenceTextFound Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "The image formula Equation cross-reference has no formula text" & _
            " [items=" & crossReferenceInventory & "]."
    End If
    If numberRange.Paragraphs.Count <> 1 Then
        Err.Raise vbObjectError + 7490, "VisualTeX", _
            "The Equation caption escaped the formula layout paragraph."
    End If
    regressionStage = "image-cross-reference-assert-layout"
    VTAssertNumberedEquationLayout _
        placeholder.Range.Duplicate, placeholder.Height, fixtureFormulaId, _
        "x^2 + y^2", "image-standard"

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
    Set sequenceField = VTFindEquationSequenceField( _
        equationRange.Paragraphs(1).Range)
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
    Set sequenceField = VTFindEquationSequenceField( _
        equationRange.Paragraphs(1).Range)
    If sequenceField.Result.Text <> previousNumberText Then
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
    Set sequenceField = VTFindEquationSequenceField( _
        equationRange.Paragraphs(1).Range)
    If sequenceField.Result.Text <> previousNumberText Then
        Err.Raise vbObjectError + 7529, "VisualTeX", _
            "Image-to-native conversion changed its Equation number."
    End If

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

Public Sub VTInitializeWordEvents()
    Set VT_WORD_EVENT_SINK = New VTWordEvents
    Set VT_WORD_EVENT_SINK.App = Word.Application
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

    ' This entry point is invoked by the native macOS double-click monitor.
    ' Do not show a modal VBA error here: a non-VisualTeX target must simply
    ' fail back to the compatibility bridge without interrupting Word.
    VTRequireWritableWordDocument
    Set selectedShape = VTVisualTeXInlineShapeAtSelection(Selection)
    If Not selectedShape Is Nothing Then
        VTWordEditInlineShape selectedShape
        Exit Sub
    End If
    VTWordEditNativeBookmark VTFindNativeFormulaBookmark(Selection.Range)
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

Public Function VTIsVisualTeXInlineShape(ByVal selectedShape As InlineShape) As Boolean
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean

    On Error GoTo InvalidShape
    If selectedShape Is Nothing Then Exit Function
    If Not VTIsEncodedMetadata(selectedShape.AlternativeText) Then Exit Function
    If Not VTTryParseFormulaReference(selectedShape.Title, formulaId, displayMode, numbered) Then Exit Function
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
    Dim field As Field
    Dim updated As Long
    Dim equationLabelName As String

    On Error GoTo Failed
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7401, "VisualTeX", "Open a Word document first."
    End If
    equationLabelName = VTNativeEquationLabelName()
    For Each field In ActiveDocument.Fields
        If VTIsNativeEquationSequenceField(field, equationLabelName) Then
            field.Update
            updated = updated + 1
        End If
    Next field
    VTShowInformation "Updated " & CStr(updated) & " native Word equation numbers."
    Exit Sub

Failed:
    VTShowError "equation numbering", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_OpenEquationCrossReference()
    Dim crossReferenceDialog As Dialog
    Dim equationLabelName As String

    On Error GoTo Failed
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7401, "VisualTeX", "Open a Word document first."
    End If
    equationLabelName = VTNativeEquationLabelName()
    Set crossReferenceDialog = Application.Dialogs(wdDialogInsertCrossReference)
    crossReferenceDialog.ReferenceType = equationLabelName
    crossReferenceDialog.ReferenceKind = wdOnlyLabelAndNumber
    crossReferenceDialog.Show
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
    If Not VTPathFileExists(VTPlaceholderImagePath()) Then
        Err.Raise vbObjectError + 7404, "VisualTeX", "The VisualTeX placeholder resource is missing. Repair the offline add-in."
    End If

    sessionId = VTNewUuidV4()
    formulaId = VTNewUuidV4()
    pendingMarker = VTPendingMarker(sessionId, formulaId)
    Set insertionRange = Selection.Range.Duplicate
    insertionRange.Collapse wdCollapseStart
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
                Exit Sub
            End If
        Else
            Set committed = VTFindCommittedInlineShape(metadata, formulaReference)
            If Not committed Is Nothing Then
                VTSetWordLatexPayload targetDocument, formulaId, latexBase64
                VTSetWordOmmlPayload targetDocument, formulaId, ommlBase64
                VTSetWordMetadataPayload targetDocument, formulaId, metadata
                VTSetWordFormulaFormat targetDocument, formulaId, displayMode, numbered
                VTDeletePendingBookmark targetDocument, sessionId
                Exit Sub
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
        If displayMode = "inline" Then
            VTPlaceCaretAfterInlineNativeEquation nativeEquationRange
        Else
            nativeEquationRange.Select
        End If
        On Error GoTo 0
        Exit Sub
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
    candidate.Select
    On Error GoTo 0
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
    Dim operationErrorNumber As Long
    Dim operationErrorDescription As String

    If formulaRange Is Nothing Or formulaRange.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7542, "VisualTeX", _
            "The numbered formula image backup target is invalid."
    End If
    Set documentObject = formulaRange.Document
    formulaStart = formulaRange.Start
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
    ByVal captionText As String) As Range

    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim insertionRange As Range
    Dim sequenceField As Field
    Dim paragraphStart As Long
    Dim equationLabelName As String
    Dim fieldAnchor As Long
    Dim operationStage As String
    Dim operationErrorNumber As Long
    Dim operationErrorDescription As String

    On Error GoTo NumberFailed
    If formulaShape Is Nothing Then
        Err.Raise vbObjectError + 7502, "VisualTeX", _
            "The numbered formula image is missing."
    End If

    operationStage = "configure-paragraph"
    Set documentObject = formulaShape.Range.Document
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    paragraphStart = paragraphRange.Start
    VTConfigureNumberedEquationParagraph paragraphRange
    equationLabelName = VTNativeEquationLabelName()

    operationStage = "insert-canonical-separator"
    Set insertionRange = formulaShape.Range.Duplicate
    insertionRange.Collapse wdCollapseEnd
    insertionRange.InsertBefore vbTab & "("
    insertionRange.Collapse wdCollapseEnd

    operationStage = "insert-caption-field"
    Set sequenceField = VTInsertRegisteredEquationCaption( _
        insertionRange, equationLabelName)
    sequenceField.Update
    fieldAnchor = VTEquationFieldStart(sequenceField)

    operationStage = "normalize-with-field"
    VTNormalizeEquationNumberLayoutWithField _
        formulaShape.Range.Duplicate, formulaShape.Height, _
        formulaId, captionText, fieldAnchor

    operationStage = "return-paragraph-range"
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

    Dim raisePoints As Single

    If formulaHeightPoints <= 0# Or formulaHeightPoints > 4096# Then
        VTEquationNumberRaisePoints = 0!
        Exit Function
    End If
    If numberFontSize <= 0! Or numberFontSize > 72! Then numberFontSize = 12!
    raisePoints = CSng(formulaHeightPoints / 2#) - numberFontSize * 0.32
    If raisePoints < 0! Then raisePoints = 0!
    If raisePoints > 48! Then raisePoints = 48!
    VTEquationNumberRaisePoints = raisePoints
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
    If InStr(1, equationLabelName, " ", vbBinaryCompare) > 0 Then
        VTEquationSequenceFieldText = """" & Replace$(equationLabelName, """", """""") & """ \* ARABIC"
    Else
        VTEquationSequenceFieldText = equationLabelName & " \* ARABIC"
    End If
End Function

Private Function VTInsertRegisteredEquationCaption( _
    ByVal insertionRange As Range, _
    ByVal equationLabelName As String) As Field

    Dim documentObject As Document
    Dim candidate As Field
    Dim match As Field
    Dim fieldParagraphRange As Range
    Dim insertionStart As Long
    Dim insertionParagraphStart As Long
    Dim sourceFieldAnchor As Long
    Dim candidateDistance As Long
    Dim bestDistance As Long
    Dim matchCount As Long
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

    ' InsertCaption is required for native Equation cross-reference
    ' registration. On Word for Mac, a collapsed Range with ExcludeLabel=True
    ' receives the registered SEQ field directly at that Range; it does not
    ' create a separate caption paragraph. Keep that native field in place.
    captionStage = "insert-caption"
    insertionRange.InsertCaption _
        Label:=wdCaptionEquation, _
        Title:="", _
        Position:=wdCaptionPositionBelow, _
        ExcludeLabel:=True

    captionStage = "find-registered-field"
    bestDistance = 2147483647
    For Each candidate In documentObject.Fields
        If VTIsNativeEquationSequenceField(candidate, equationLabelName) Then
            candidateDistance = Abs( _
                VTEquationFieldStart(candidate) - insertionStart)
            If candidateDistance <= 64 Then
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
        Err.Raise vbObjectError + 7425, "VisualTeX", _
            "Word did not register the Equation caption for cross-reference."
    End If

    captionStage = "verify-inline-field"
    sourceFieldAnchor = VTEquationFieldStart(match)
    Set fieldParagraphRange = match.Result.Paragraphs(1).Range.Duplicate
    If fieldParagraphRange.Start <> insertionParagraphStart Or _
       Abs(sourceFieldAnchor - insertionStart) > 16 Then
        Err.Raise vbObjectError + 7537, "VisualTeX", _
            "Word inserted the Equation caption outside the formula paragraph."
    End If

    Set VTInsertRegisteredEquationCaption = match
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
    Dim equationLabelName As String
    Dim candidateStart As Long
    Dim candidateDistance As Long
    Dim bestDistance As Long
    Dim matchCount As Long

    If documentObject Is Nothing Or expectedStart < 0 Or _
       maximumDistance < 0 Then
        Err.Raise vbObjectError + 7536, "VisualTeX", _
            "The Equation field resolver received an invalid target."
    End If
    equationLabelName = VTNativeEquationLabelName()
    bestDistance = 2147483647
    For Each candidate In documentObject.Fields
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

Private Function VTFindEquationSequenceField( _
    ByVal paragraphRange As Range) As Field

    Dim candidate As Field
    Dim match As Field
    Dim matchCount As Long
    Dim equationLabelName As String

    If paragraphRange Is Nothing Then Exit Function
    equationLabelName = VTNativeEquationLabelName()
    For Each candidate In paragraphRange.Fields
        If VTIsNativeEquationSequenceField(candidate, equationLabelName) Then
            matchCount = matchCount + 1
            Set match = candidate
        End If
    Next candidate
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
    paragraphRange.Style = wdStyleCaption
    With paragraphRange.ParagraphFormat
        .Alignment = wdAlignParagraphLeft
        .LeftIndent = 0!
        .RightIndent = 0!
        .FirstLineIndent = 0!
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

    Dim bookmarkName As String
    Dim oldCaptionRange As Range
    Dim closingRange As Range
    Dim captionRange As Range
    Dim fieldAnchor As Long
    Dim fieldEnd As Long
    Dim captionStart As Long
    Dim captionValue As String

    If documentObject Is Nothing Or sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7495, "VisualTeX", _
            "The Equation caption target is missing."
    End If
    fieldAnchor = VTEquationFieldStart(sequenceField)
    bookmarkName = VTEquationCaptionBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        Set oldCaptionRange = documentObject.Bookmarks(bookmarkName).Range.Duplicate
        documentObject.Bookmarks(bookmarkName).Delete
        oldCaptionRange.Delete
    End If

    Set sequenceField = VTResolveEquationSequenceFieldNear( _
        documentObject, fieldAnchor, 256)
    fieldEnd = VTEquationFieldEnd(sequenceField)
    If fieldEnd >= documentObject.Content.End Or _
       documentObject.Range(fieldEnd, fieldEnd + 1).Text <> ")" Then
        Set closingRange = documentObject.Range( _
            Start:=fieldEnd, End:=fieldEnd)
        closingRange.Text = ")"
    End If
    captionStart = fieldEnd + 1
    captionValue = " " & captionText
    Set captionRange = documentObject.Range( _
        Start:=captionStart, End:=captionStart)
    captionRange.Text = captionValue
    Set captionRange = documentObject.Range( _
        Start:=captionStart, End:=captionStart + Len(captionValue))
    With captionRange.Font
        .Hidden = False
        .Color = wdColorWhite
        .Size = 1!
        .Scaling = 1
        .Spacing = -1!
    End With
    documentObject.Bookmarks.Add Name:=bookmarkName, Range:=captionRange
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
    Dim openingRange As Range
    Dim resultRange As Range
    Dim closingRange As Range
    Dim fieldStart As Long
    Dim fieldEnd As Long
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
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    paragraphStart = paragraphRange.Start
    VTConfigureNumberedEquationParagraph paragraphRange

    operationStage = "normalize-center-prefix"
    formulaWasImage = formulaRange.InlineShapes.Count = 1
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
    numberRaisePoints = VTEquationNumberRaisePoints( _
        renderedHeightPoints, preferredSize)
    openingRange.Font.Position = CLng(numberRaisePoints)
    resultRange.Font.Position = CLng(numberRaisePoints)
    closingRange.Font.Position = CLng(numberRaisePoints)

    operationStage = "format-formula"
    formulaRange.Font.Position = 0
    If formulaRange.OMaths.Count = 1 Then
        formulaRange.OMaths(1).Type = wdOMathInline
        formulaRange.Font.Size = preferredSize
    End If

    operationStage = "verify-layout-tabs"
    Set sequenceField = VTResolveEquationSequenceFieldNear( _
        documentObject, fieldAnchor, 256)
    fieldAnchor = VTEquationFieldStart(sequenceField)
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    If paragraphRange.Document.Range( _
        paragraphRange.Start, paragraphRange.Start + 1).Text <> vbTab Then
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

Private Sub VTAssertNumberedEquationLayout( _
    ByVal formulaRange As Range, _
    ByVal renderedHeightPoints As Double, _
    ByVal formulaId As String, _
    ByVal expectedCaptionText As String, _
    ByVal assertionName As String)

    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim sequenceField As Field
    Dim openingRange As Range
    Dim resultRange As Range
    Dim closingRange As Range
    Dim captionRange As Range
    Dim formulaStartProbe As Range
    Dim formulaEndProbe As Range
    Dim numberEndProbe As Range
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
    Dim bookmarkName As String

    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7505, "VisualTeX", _
            assertionName & ": formula Range is missing."
    End If
    Set documentObject = formulaRange.Document
    Set paragraphRange = formulaRange.Paragraphs(1).Range.Duplicate
    Set sequenceField = VTFindEquationSequenceField(paragraphRange)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7506, "VisualTeX", _
            assertionName & ": Equation SEQ field is missing."
    End If

    textWidth = VTEquationLayoutWidth(paragraphRange)
    If paragraphRange.ParagraphFormat.Alignment <> wdAlignParagraphLeft Then
        Err.Raise vbObjectError + 7507, "VisualTeX", _
            assertionName & ": numbered paragraph is not left-aligned."
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
    fieldStart = VTEquationFieldStart(sequenceField)
    fieldEnd = VTEquationFieldEnd(sequenceField)
    If fieldStart < 2 Or _
       documentObject.Range( _
           fieldStart - 2, _
           fieldStart - 1).Text <> vbTab Then
        Err.Raise vbObjectError + 7512, "VisualTeX", _
            assertionName & ": Equation number is not anchored by the right tab."
    End If

    Set openingRange = documentObject.Range( _
        Start:=fieldStart - 1, _
        End:=fieldStart)
    Set closingRange = documentObject.Range( _
        Start:=fieldEnd, _
        End:=fieldEnd + 1)
    If openingRange.Text <> "(" Or closingRange.Text <> ")" Then
        Err.Raise vbObjectError + 7513, "VisualTeX", _
            assertionName & ": Equation number parentheses are incomplete."
    End If
    Set resultRange = sequenceField.Result.Duplicate
    preferredSize = resultRange.Font.Size
    If preferredSize <= 0! Or preferredSize > 72! Or _
       openingRange.Font.Size <> preferredSize Or _
       closingRange.Font.Size <> preferredSize Then
        Err.Raise vbObjectError + 7514, "VisualTeX", _
            assertionName & ": Equation number font size is invalid."
    End If
    expectedRaise = CLng(VTEquationNumberRaisePoints( _
        renderedHeightPoints, preferredSize))
    If openingRange.Font.Position <> expectedRaise Or _
       resultRange.Font.Position <> expectedRaise Or _
       closingRange.Font.Position <> expectedRaise Then
        Err.Raise vbObjectError + 7515, "VisualTeX", _
            assertionName & ": Equation number is not vertically centered."
    End If

    bookmarkName = VTEquationCaptionBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(bookmarkName) Then
        Err.Raise vbObjectError + 7516, "VisualTeX", _
            assertionName & ": searchable Equation caption text is missing."
    End If
    Set captionRange = documentObject.Bookmarks(bookmarkName).Range.Duplicate
    If InStr(1, captionRange.Text, expectedCaptionText, vbTextCompare) = 0 Or _
       captionRange.Font.Hidden <> False Or _
       captionRange.Font.Color <> wdColorWhite Or _
       captionRange.Font.Size <> 1! Or _
       captionRange.Font.Scaling <> 1 Then
        Err.Raise vbObjectError + 7517, "VisualTeX", _
            assertionName & ": searchable Equation caption text is invalid" & _
            " [text=" & Replace$(Replace$(captionRange.Text, vbTab, "<TAB>"), vbCr, "<CR>") & _
            "; expected=" & expectedCaptionText & _
            "; hidden=" & CStr(captionRange.Font.Hidden) & _
            "; color=" & CStr(captionRange.Font.Color) & _
            "; size=" & CStr(captionRange.Font.Size) & _
            "; scaling=" & CStr(captionRange.Font.Scaling) & _
            "; range=" & CStr(captionRange.Start) & "-" & _
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
            ' Word centers OMath by its full typographic box, but Range.End
            ' excludes the invisible terminal math zone from its horizontal
            ' position. Half one math font size covers that measurement gap;
            ' image formulas retain the strict three-point geometry threshold.
            If centerTolerance < nativeFontSize / 2! Then
                centerTolerance = nativeFontSize / 2!
            End If
        End If
    End If
    If Abs(formulaCenterPosition - textWidth / 2!) > centerTolerance Then
        Err.Raise vbObjectError + 7519, "VisualTeX", _
            assertionName & ": formula is not geometrically centered" & _
            " [range=" & CStr(formulaRange.Start) & "-" & _
            CStr(formulaRange.End) & _
            "; startPosition=" & CStr(formulaStartPosition) & _
            "; endPosition=" & CStr(formulaEndPosition) & _
            "; center=" & CStr(formulaCenterPosition) & _
            "; target=" & CStr(textWidth / 2!) & _
            "; tolerance=" & CStr(centerTolerance) & _
            "; textWidth=" & CStr(textWidth) & "]."
    End If
    If Abs(numberEndPosition - (textWidth - 1!)) > 3! Then
        Err.Raise vbObjectError + 7520, "VisualTeX", _
            assertionName & ": Equation number is not at the right text boundary" & _
            " [numberEnd=" & CStr(numberEndPosition) & _
            "; target=" & CStr(textWidth - 1!) & _
            "; delta=" & CStr(numberEndPosition - (textWidth - 1!)) & _
            "; field=" & CStr(fieldStart) & "-" & CStr(fieldEnd) & _
            "; textWidth=" & CStr(textWidth) & "]."
    End If
    formulaLine = formulaStartProbe.Information(wdFirstCharacterLineNumber)
    numberLine = numberEndProbe.Information(wdFirstCharacterLineNumber)
    If formulaLine <= 0 Or numberLine <= 0 Or formulaLine <> numberLine Then
        Err.Raise vbObjectError + 7521, "VisualTeX", _
            assertionName & ": formula and Equation number are not on the same line."
    End If
End Sub

Private Function VTEnsureImageEquationNumber( _
    ByRef formulaShape As InlineShape, _
    ByVal renderedHeightPoints As Double, _
    ByVal formulaId As String, _
    ByVal captionText As String, _
    ByRef numberCreated As Boolean) As Range

    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim sequenceField As Field
    Dim paragraphStart As Long

    If formulaShape Is Nothing Then
        Err.Raise vbObjectError + 7502, "VisualTeX", _
            "The numbered formula image is missing."
    End If
    Set documentObject = formulaShape.Range.Document
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    paragraphStart = paragraphRange.Start
    Set sequenceField = VTFindEquationSequenceField(paragraphRange)
    If sequenceField Is Nothing Then
        numberCreated = True
        Set VTEnsureImageEquationNumber = VTInsertEquationNumber( _
            formulaShape, formulaId, captionText)
    Else
        numberCreated = False
        VTNormalizeEquationNumberLayout _
            formulaShape.Range.Duplicate, renderedHeightPoints, _
            formulaId, captionText
        Set formulaShape = VTResolveImageFormulaInParagraph( _
            documentObject, paragraphStart)
        Set VTEnsureImageEquationNumber = _
            formulaShape.Range.Paragraphs(1).Range.Duplicate
    End If
End Function

Private Function VTEnsureNativeEquationNumber( _
    ByVal equationRange As Range, _
    ByVal renderedHeightPoints As Double, _
    ByVal formulaId As String, _
    ByVal captionText As String, _
    ByRef numberCreated As Boolean) As Range

    Dim sequenceField As Field

    If equationRange Is Nothing Or equationRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7470, "VisualTeX", _
            "The native equation number target is missing."
    End If
    equationRange.OMaths(1).Type = wdOMathInline
    Set equationRange = equationRange.OMaths(1).Range.Duplicate
    Set sequenceField = VTFindEquationSequenceField( _
        equationRange.Paragraphs(1).Range)
    If sequenceField Is Nothing Then
        numberCreated = True
        Set VTEnsureNativeEquationNumber = VTInsertNativeEquationNumber( _
            equationRange, renderedHeightPoints, formulaId, captionText)
    Else
        numberCreated = False
        VTNormalizeEquationNumberLayout _
            equationRange, renderedHeightPoints, formulaId, captionText
        Set VTEnsureNativeEquationNumber = _
            equationRange.Paragraphs(1).Range.Duplicate
    End If
End Function

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
    Set paragraphRange = exactEquationRange.Paragraphs(1).Range.Duplicate
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
    preferredSize = VTPreferredEquationFontSize( _
        targetRange, _
        displaySizing Or displayMode = "block")
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
    Dim numberCreated As Boolean

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
    VTSetNativeFormulaBookmark targetDocument, equationRange, formulaId

    sourceBackupDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set sourceBackupDocument = Nothing
    On Error Resume Next
    If displayMode = "inline" Then
        VTPlaceCaretAfterInlineNativeEquation equationRange
    Else
        equationRange.Select
    End If
    On Error GoTo 0
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

Private Sub VTWriteWordHealth()
    Dim statusPath As String
    Dim payload As String
    statusPath = VTApplicationSupportRoot() & VT_WORD_STATUS_FILE
    payload = "{" & _
        """loaded"":true," & _
        """pluginVersion"":" & VTJsonString(VT_PLUGIN_VERSION) & "," & _
        """host"":""word""," & _
        """timestamp"":" & VTJsonString(Format$(Now, "yyyy-mm-dd\Thh:nn:ss")) & _
        "}"
    VTWriteTextAtomic statusPath, payload
End Sub
