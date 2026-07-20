Attribute VB_Name = "VTWordAdapter"
Option Explicit

Private Const VT_WORD_HOST As String = "word"
Private Const VT_WORD_STATUS_FILE As String = "/OfficePluginStatus/word.json"
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
    Set insertionRange = VTInsertEquationNumber( _
        displayFormula, imageFormulaId, "workflow image formula")
    Set imageTable = insertionRange.Tables(1)
    VTVerifyEquationNumberFieldIntegrity imageTable, imageFormulaId, 1

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
    Set imageTable = testDocument.Bookmarks( _
        VTEquationNumberBookmarkName(imageFormulaId)).Range.Tables(1)
    Set nativeTable = testDocument.Bookmarks( _
        VTEquationNumberBookmarkName(nativeFormulaId)).Range.Tables(1)
    VTVerifyEquationNumberFieldIntegrity imageTable, imageFormulaId, 1
    VTVerifyEquationNumberFieldIntegrity nativeTable, nativeFormulaId, 2

    regressionStage = "native-display-continuation"
    Set nativeEquationRange = nativeTable.Cell(1, 2).Range.OMaths(1).Range.Duplicate
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
    Set orphanRange = nativeTable.Cell(1, 2).Range.Duplicate
    nativeTable.Cell(1, 2).Range.OMaths(1).Range.Delete
    VTCleanupOrphanedNumberedDisplaySelection orphanRange
    Set imageTable = testDocument.Bookmarks( _
        VTEquationNumberBookmarkName(imageFormulaId)).Range.Tables(1)
    VTReconcileEquationNumbers testDocument
    VTVerifyEquationNumberFieldIntegrity imageTable, imageFormulaId, 1

    regressionStage = "delete-image-numbered-display"
    Set orphanRange = imageTable.Cell(1, 2).Range.Duplicate
    imageTable.Cell(1, 2).Range.InlineShapes(1).Delete
    VTCleanupOrphanedNumberedDisplaySelection orphanRange
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
    Set VTRegressionCreateNumberedImage = _
        numberRange.Tables(1).Cell(1, 2).Range.InlineShapes(1)
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
    Set equationRange = _
        numberRange.Tables(1).Cell(1, 2).Range.OMaths(1).Range.Duplicate
    VTSetNativeFormulaBookmark documentObject, equationRange, formulaId
    Set VTRegressionCreateNumberedNative = equationRange.Duplicate
End Function

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
    Set layoutTable = testDocument.Bookmarks( _
        VTEquationNumberBookmarkName(nativeFormulaId2)).Range.Tables(1)
    layoutTable.Cell(1, 2).Range.OMaths(1).Range.Delete
    Set layoutTable = testDocument.Bookmarks( _
        VTEquationNumberBookmarkName(imageFormulaId4)).Range.Tables(1)
    layoutTable.Cell(1, 2).Range.InlineShapes(1).Delete

    regressionStage = "prune-and-renumber-after-deletion"
    liveCount = VTPruneOrphanedEquationNumberScaffolds(testDocument)
    VTReconcileEquationNumbers testDocument
    If liveCount <> 2 Or testDocument.Tables.Count <> 2 Then
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
    Set layoutTable = testDocument.Bookmarks( _
        VTEquationNumberBookmarkName(imageFormulaId1)).Range.Tables(1)
    VTVerifyEquationNumberFieldIntegrity layoutTable, imageFormulaId1, 1
    Set layoutTable = testDocument.Bookmarks( _
        VTEquationNumberBookmarkName(nativeFormulaId3)).Range.Tables(1)
    VTVerifyEquationNumberFieldIntegrity layoutTable, nativeFormulaId3, 2

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
    Set layoutTable = testDocument.Bookmarks( _
        VTEquationNumberBookmarkName(firstFormulaId)).Range.Tables(1)
    layoutTable.Cell(1, 2).Range.InlineShapes(1).Delete
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
    Set layoutTable = testDocument.Bookmarks( _
        VTEquationNumberBookmarkName(referencedFormulaId)).Range.Tables(1)
    Set referencedFormula = _
        layoutTable.Cell(1, 2).Range.InlineShapes(1)
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
    Dim candidateField As Field
    Dim builtInReferenceField As Field
    Dim nativeItems As Variant
    Dim fixtureRoot As String
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String
    Dim bodyReferenceCount As Long
    Dim nativeSequenceCount As Long
    Dim latestBodyReferenceStart As Long
    Dim candidateFieldStart As Long

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

    regressionStage = "verify-native-caption-inventory"
    nativeItems = testDocument.GetCrossReferenceItems(wdCaptionEquation)
    If VTVariantArrayCount(nativeItems) <> 2 Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "Word did not expose exactly two VisualTeX Equation captions."
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

    regressionStage = "insert-word-built-in-reference"
    Set insertionRange = VTAppendRegressionParagraph(testDocument)
    insertionRange.Select
    Selection.InsertCrossReference _
        ReferenceType:=wdCaptionEquation, _
        ReferenceKind:=wdOnlyLabelAndNumber, _
        ReferenceItem:=2, _
        InsertAsHyperlink:=True, _
        IncludePosition:=False

    regressionStage = "verify-word-built-in-initial-visibility"
    Set builtInReferenceField = Nothing
    latestBodyReferenceStart = -1
    For Each candidateField In testDocument.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            candidateFieldStart = VTEquationFieldStart(candidateField)
            If candidateFieldStart > latestBodyReferenceStart Then
                latestBodyReferenceStart = candidateFieldStart
                Set builtInReferenceField = candidateField
            End If
        End If
    Next candidateField
    If builtInReferenceField Is Nothing Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "Word did not create the built-in Equation reference."
    End If
    If Trim$(builtInReferenceField.Result.Text) <> "2" Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "The initial Word built-in Equation reference result is not 2" & _
            " [result=" & builtInReferenceField.Result.Text & "]."
    End If
    VTAssertBodyEquationReferenceVisible _
        testDocument, builtInReferenceField, _
        "initial Word built-in Equation reference"

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
    Set layoutTable = testDocument.Bookmarks( _
        VTEquationNumberBookmarkName(firstFormulaId)).Range.Tables(1)
    layoutTable.Cell(1, 2).Range.InlineShapes(1).Delete
    VTPruneOrphanedEquationNumberScaffolds testDocument
    VTReconcileEquationNumbers testDocument
    nativeItems = testDocument.GetCrossReferenceItems(wdCaptionEquation)
    If VTVariantArrayCount(nativeItems) <> 1 Then
        Err.Raise vbObjectError + 7555, "VisualTeX", _
            "Deleting the first formula did not leave exactly one native" & _
            " Equation caption."
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
        "initialVisualTeXReference=2" & vbLf & _
        "initialWordBuiltInReference=2" & vbLf & _
        "renumberedVisualTeXReference=1" & vbLf & _
        "renumberedWordBuiltInReference=1" & vbLf
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
    If numberRange.Tables.Count <> 1 Or _
       numberRange.Tables(1).Rows.Count <> 1 Or _
       numberRange.Tables(1).Columns.Count <> 3 Then
        Err.Raise vbObjectError + 7490, "VisualTeX", _
            "The Equation number escaped the stable display table."
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
            "The native Equation REF is not the exact parenthesized number" & _
            " [code=" & referenceField.Code.Text & _
            "; result=" & referenceField.Result.Text & _
            "; text=" & referenceResult & "]."
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
    Set sequenceField = VTFindEquationSequenceField( _
        equationRange.Paragraphs(1).Range)
    If sequenceField.Result.Text <> previousNumberText Then
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
    Optional ByVal referenceResultText As String = "") As String

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
    formulaIds = VTValidNumberedFormulaIds(documentObject)
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

    If documentObject Is Nothing Or sequenceField Is Nothing Then Exit Function
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

Private Function VTValidNumberedFormulaIds( _
    ByVal documentObject As Document) As Variant

    Dim layoutTable As Table
    Dim seenIds As New Collection
    Dim formulaId As String
    Dim ids() As String
    Dim itemCount As Long

    If documentObject Is Nothing Then Exit Function
    For Each layoutTable In documentObject.Tables
        formulaId = VTFormulaIdForValidNumberedTable(layoutTable)
        If Len(formulaId) > 0 Then
            If VTCollectionContainsText(seenIds, formulaId) Then
                Err.Raise vbObjectError + 7553, "VisualTeX", _
                    "Two numbered VisualTeX tables use the same formula id" & _
                    " [formulaId=" & formulaId & "]."
            End If
            VTCollectionAddUniqueText seenIds, formulaId
            itemCount = itemCount + 1
            ReDim Preserve ids(1 To itemCount)
            ids(itemCount) = formulaId
        End If
    Next layoutTable
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
       paragraphRange.Fields.Count <> 1 Then Exit Function
    Set candidateField = paragraphRange.Fields(1)
    VTHelperParagraphOwnsNativeEquationSequence = _
        VTIsNativeEquationSequenceField( _
            candidateField, VTNativeEquationLabelName())
End Function

Private Sub VTDeleteEquationNumberScaffold( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    Optional ByVal deleteTable As Boolean = True)

    Dim numberBookmarkName As String
    Dim sequenceBookmarkName As String
    Dim captionBookmarkName As String
    Dim nativeBookmarkName As String
    Dim helperParagraph As Range
    Dim layoutTable As Table
    Dim sequenceField As Field

    If documentObject Is Nothing Or Not VTIsCanonicalUuid(formulaId) Then
        Exit Sub
    End If
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    sequenceBookmarkName = VTEquationSequenceNumberBookmarkName(formulaId)
    captionBookmarkName = VTEquationCaptionBookmarkName(formulaId)
    nativeBookmarkName = VTNativeFormulaBookmarkName(formulaId)

    If documentObject.Bookmarks.Exists(numberBookmarkName) Then
        On Error Resume Next
        If documentObject.Bookmarks(numberBookmarkName).Range.Information( _
           wdWithInTable) Then
            Set layoutTable = _
                documentObject.Bookmarks(numberBookmarkName).Range.Tables(1)
        End If
        On Error GoTo 0
    End If
    If documentObject.Bookmarks.Exists(captionBookmarkName) Then
        Set helperParagraph = documentObject.Bookmarks( _
            captionBookmarkName).Range.Duplicate
    ElseIf documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        Set sequenceField = VTEquationSequenceFieldForBookmark( _
            documentObject, sequenceBookmarkName)
        If Not sequenceField Is Nothing Then
            Set helperParagraph = _
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

    If Not helperParagraph Is Nothing Then
        If VTHelperParagraphOwnsNativeEquationSequence( _
           helperParagraph) Then
            helperParagraph.Delete
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

    VTPruneUnbookmarkedEmptyNumberTables documentObject
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
    Dim numberBookmarkName As String
    Dim suffixText As String
    Dim formulaId As String
    Dim tableStart As Long

    If selectedRange Is Nothing Or _
       VTWordInternalMutationActive() Then Exit Sub
    Set layoutTable = VTNumberedDisplayTableNearRange(selectedRange)
    If layoutTable Is Nothing Then Exit Sub
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
    tableStart = layoutTable.Range.Start

    VTDeleteEquationNumberScaffold documentObject, formulaId, True
    VTReconcileEquationNumbers documentObject

    On Error Resume Next
    If tableStart > documentObject.Content.End - 1 Then
        tableStart = documentObject.Content.End - 1
    End If
    If tableStart < 0 Then tableStart = 0
    Set caretRange = documentObject.Range( _
        Start:=tableStart, End:=tableStart)
    Set paragraphRange = caretRange.Paragraphs(1).Range.Duplicate
    VTNormalizePlainWordParagraph paragraphRange
    If documentObject Is ActiveDocument Then caretRange.Select
    On Error GoTo 0
End Sub

Private Function VTAnyOpenDocumentHasEquationNumbers() As Boolean
    Dim documentObject As Document
    Dim candidateBookmark As Bookmark

    For Each documentObject In Documents
        For Each candidateBookmark In documentObject.Bookmarks
            If Left$(candidateBookmark.Name, _
               Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
               VT_WORD_NUMBER_BOOKMARK_PREFIX Then
                VTAnyOpenDocumentHasEquationNumbers = True
                Exit Function
            End If
        Next candidateBookmark
    Next documentObject
End Function

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
        VTCleanupOrphanedNumberedDisplaySelection Selection.Range
        VTNormalizeBodyEquationReferenceVisibility ActiveDocument
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
    Dim nativeItems As Variant
    Dim itemIndex As Long
    Dim itemCount As Long
    Dim nativeItemCount As Long
    Dim nativeReferenceItem As Long
    Dim sequenceBookmarkName As String
    Dim numberText As String
    Dim items() As String

    If documentObject Is Nothing Then Exit Function
    VTPruneOrphanedEquationNumberScaffolds documentObject
    VTReconcileEquationNumbers documentObject
    formulaIds = VTValidNumberedFormulaIds(documentObject)
    itemCount = VTVariantArrayCount(formulaIds)
    If itemCount <= 0 Then Exit Function

    nativeItems = documentObject.GetCrossReferenceItems(wdCaptionEquation)
    nativeItemCount = VTVariantArrayCount(nativeItems)
    If nativeItemCount <= 0 Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "Word exposes no native Equation cross-reference items."
    End If

    ReDim items(1 To itemCount)
    For itemIndex = 1 To itemCount
        nativeReferenceItem = VTNativeEquationReferenceItemForFormula( _
            documentObject, CStr(formulaIds(itemIndex)))
        If nativeReferenceItem < 1 Or _
           nativeReferenceItem > nativeItemCount Then
            Err.Raise vbObjectError + 7547, "VisualTeX", _
                "A VisualTeX formula is missing from Word's native Equation" & _
                " cross-reference inventory."
        End If
        sequenceBookmarkName = VTEquationSequenceNumberBookmarkName( _
            CStr(formulaIds(itemIndex)))
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
    Dim nativeItems As Variant
    Dim formulaId As String
    Dim sequenceBookmarkName As String
    Dim expectedNumber As String
    Dim referenceField As Field
    Dim insertedRange As Range
    Dim insertionStart As Long
    Dim insertionEnd As Long
    Dim itemCount As Long
    Dim nativeItemCount As Long
    Dim nativeReferenceItem As Long

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

    nativeItems = documentObject.GetCrossReferenceItems(wdCaptionEquation)
    nativeItemCount = VTVariantArrayCount(nativeItems)
    nativeReferenceItem = VTNativeEquationReferenceItemForFormula( _
        documentObject, formulaId)
    If nativeReferenceItem < 1 Or _
       nativeReferenceItem > nativeItemCount Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "Word cannot resolve the selected native Equation caption item."
    End If

    insertionStart = targetRange.Start
    targetRange.Select
    Selection.Collapse wdCollapseStart
    Selection.TypeText Text:="("
    Selection.InsertCrossReference _
        ReferenceType:=wdCaptionEquation, _
        ReferenceKind:=wdEntireCaption, _
        ReferenceItem:=nativeReferenceItem, _
        InsertAsHyperlink:=True, _
        IncludePosition:=False
    Selection.Collapse wdCollapseEnd
    Selection.TypeText Text:=")"
    insertionEnd = Selection.Range.Start
    Set insertedRange = documentObject.Range( _
        Start:=insertionStart, End:=insertionEnd)
    If insertedRange.Fields.Count <> 1 Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "Word did not create exactly one native Equation REF field."
    End If
    Set referenceField = insertedRange.Fields(1)
    referenceField.Update
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

    ' A display formula must never reuse a body paragraph that already contains
    ' text, an inline formula, an image, or native OMath. The old path inserted a
    ' one-pixel placeholder at the caret and later cleared that entire paragraph
    ' while building the numbered table, erasing any earlier inline formula on
    ' the same line. Split the paragraph first and return a dedicated empty line.
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
        beforeRange.InlineShapes.Count > 0 Or _
        beforeRange.OMaths.Count > 0 Or _
        VTWordRangeHasMeaningfulText(beforeRange)
    afterOccupied = _
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
    VTNormalizePlainWordParagraph targetParagraph
    Set VTPrepareWordCreateInsertionRange = insertionRange
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
    Dim insertionRange As Range
    Dim helperParagraph As Range
    Dim helperStart As Long

    If layoutTable Is Nothing Then
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "The Equation helper paragraph requires a numbered table."
    End If
    Set documentObject = layoutTable.Range.Document
    helperStart = layoutTable.Range.End
    Set insertionRange = documentObject.Range( _
        Start:=helperStart, End:=helperStart)

    ' Every numbered formula owns a distinct helper paragraph. Never clear or
    ' reuse the paragraph that already follows the table: it can be body text,
    ' another formula, or another formula's native SEQ helper.
    insertionRange.Text = vbCr
    Set helperParagraph = documentObject.Range( _
        Start:=helperStart, End:=helperStart).Paragraphs(1).Range.Duplicate
    If helperParagraph.Information(wdWithInTable) Or _
       helperParagraph.Fields.Count <> 0 Or _
       helperParagraph.InlineShapes.Count <> 0 Or _
       helperParagraph.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(helperParagraph) Then
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "Word did not create a unique empty Equation helper paragraph."
    End If
    Set VTInsertDedicatedEquationHelperParagraph = helperParagraph
End Function

Private Function VTEnsurePlainContinuationParagraph( _
    ByVal sourceParagraph As Range) As Range

    Dim documentObject As Document
    Dim insertionRange As Range
    Dim continuationParagraph As Range
    Dim continuationStart As Long

    If sourceParagraph Is Nothing Then
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "The display continuation source paragraph is missing."
    End If
    Set documentObject = sourceParagraph.Document
    continuationStart = sourceParagraph.End
    If continuationStart >= documentObject.Content.End Then
        documentObject.Content.InsertAfter vbCr
    End If

    Set continuationParagraph = documentObject.Range( _
        Start:=continuationStart, End:=continuationStart).Paragraphs(1).Range.Duplicate
    If continuationParagraph.Information(wdWithInTable) Or _
       continuationParagraph.Fields.Count > 0 Or _
       continuationParagraph.InlineShapes.Count > 0 Or _
       continuationParagraph.OMaths.Count > 0 Then
        Set insertionRange = documentObject.Range( _
            Start:=continuationStart, End:=continuationStart)
        insertionRange.Text = vbCr
        Set continuationParagraph = documentObject.Range( _
            Start:=continuationStart, End:=continuationStart).Paragraphs(1).Range.Duplicate
    End If
    If continuationParagraph.Information(wdWithInTable) Or _
       continuationParagraph.Fields.Count <> 0 Or _
       continuationParagraph.InlineShapes.Count <> 0 Or _
       continuationParagraph.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7553, "VisualTeX", _
            "Word did not expose a plain paragraph after the display formula."
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
        If displayMode = "inline" Then
            VTPlaceCaretAfterInlineNativeEquation nativeEquationRange
        ElseIf mode = "create" Then
            VTPlaceCaretAfterDisplayFormula nativeEquationRange, formulaId
        Else
            nativeEquationRange.Select
        End If
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
    If displayMode = "block" And mode = "create" Then
        VTPlaceCaretAfterDisplayFormula candidate.Range, formulaId
    Else
        candidate.Select
    End If

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

    Dim numberCreated As Boolean

    Set VTInsertEquationNumber = VTEnsureImageEquationNumber( _
        formulaShape, formulaShape.Height, formulaId, captionText, _
        numberCreated)
End Function

Private Function VTEquationNumberRaisePoints( _
    ByVal formulaHeightPoints As Double, _
    ByVal numberFontSize As Single) As Single

    ' Never turn rendered image height into Font.Position. A raised number
    ' expands Word's line box above the formula and destabilizes older rows.
    VTEquationNumberRaisePoints = 0!
End Function

Private Function VTNativeEquationNumberRaisePoints( _
    ByVal formulaRange As Range, _
    ByVal numberFontSize As Single) As Single

    ' The number is part of Word's native display equation array, so Word
    ' aligns it to the mathematical axis. Any Font.Position correction would
    ' inflate the line box and recreate the large blank area above the OMath.
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
    ' invalidates Word's native cross-reference target.
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
    Dim fieldAnchor As Long
    Dim equationLabelName As String

    If documentObject Is Nothing Or sequenceField Is Nothing Or _
       Len(sequenceBookmarkName) = 0 Or sequenceOrdinal < 1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The native Equation number target is missing."
    End If
    suffixText = Mid$(sequenceBookmarkName, _
        Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX) + 1)
    numberBookmarkName = VT_WORD_NUMBER_BOOKMARK_PREFIX & suffixText
    captionBookmarkName = VT_WORD_CAPTION_BOOKMARK_PREFIX & suffixText

    ' A structural edit near a Word field can stale its COM wrapper. Resolve
    ' it again from VT_N_, update it, and rebuild VT_N_ before reading text.
    Set sequenceField = VTEquationSequenceFieldForBookmark( _
        documentObject, sequenceBookmarkName)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The Equation number target cannot resolve its native SEQ field."
    End If
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

    Set sequenceField = VTEquationSequenceFieldForBookmark( _
        documentObject, sequenceBookmarkName)
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

    If documentObject Is Nothing Or numberRange Is Nothing Or _
       numberRange.End <= numberRange.Start Then
        Err.Raise vbObjectError + 7546, "VisualTeX", _
            "The exact Equation number Bookmark target is missing."
    End If
    bookmarkName = VTEquationNumberBookmarkName(formulaId)
    If documentObject.Bookmarks.Exists(bookmarkName) Then
        documentObject.Bookmarks(bookmarkName).Delete
    End If
    documentObject.Bookmarks.Add name:=bookmarkName, Range:=numberRange
End Sub

Private Sub VTDeleteEquationCaptionText( _
    ByVal documentObject As Document, _
    ByVal formulaId As String)

    Dim bookmarkName As String
    Dim captionRange As Range

    If documentObject Is Nothing Then Exit Sub
    bookmarkName = VTEquationCaptionBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(bookmarkName) Then Exit Sub
    Set captionRange = documentObject.Bookmarks(bookmarkName).Range.Duplicate
    documentObject.Bookmarks(bookmarkName).Delete
    If captionRange.End > captionRange.Start Then captionRange.Delete
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
    paragraphRange.Style = wdStyleCaption
    With paragraphRange.ParagraphFormat
        .Alignment = wdAlignParagraphLeft
        .LeftIndent = 0!
        .RightIndent = 0!
        .FirstLineIndent = 0!
        .SpaceBefore = 0!
        .SpaceAfter = 0!
        .LineSpacingRule = wdLineSpaceSingle
        .KeepWithNext = False
        .KeepTogether = True
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
        formulaRange.OMaths(1).Type = wdOMathDisplay
        formulaRange.OMaths(1).Justification = wdOMathJcCenterGroup
        formulaRange.Font.Size = VTPreferredNativeDisplayFontSize(formulaRange)
        Set formulaRange = VTResolveNativeEquationRange( _
            documentObject, paragraphStart, 512)
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
    If documentObject.Bookmarks.Exists( _
       VTEquationCaptionBookmarkName(formulaId)) Then
        centerTabStart = documentObject.Bookmarks( _
            VTEquationCaptionBookmarkName(formulaId)).Range.End
    End If
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
    Dim equationArrayMarker As Range
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
    Dim nativeNumberInMath As Boolean

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
    nativeNumberInMath = formulaRange.OMaths.Count = 1 And _
        fieldStart > formulaRange.Start And fieldEnd < formulaRange.End
    If formulaRange.OMaths.Count = 1 Then
        If Not nativeNumberInMath Then
            Err.Raise vbObjectError + 7512, "VisualTeX", _
                assertionName & _
                ": native Equation number is outside the display OMath."
        End If
        Set equationArrayMarker = documentObject.Range( _
            formulaRange.Start, fieldStart)
        With equationArrayMarker.Find
            .ClearFormatting
            .Text = "#"
            .Forward = False
            .Wrap = wdFindStop
            .Format = False
        End With
        If Not equationArrayMarker.Find.Execute Then
            Err.Raise vbObjectError + 7512, "VisualTeX", _
                assertionName & _
                ": native Equation array marker is missing."
        End If
    ElseIf fieldStart < 2 Or _
           documentObject.Range( _
               fieldStart - 2, _
               fieldStart - 1).Text <> vbTab Then
            Err.Raise vbObjectError + 7512, "VisualTeX", _
                assertionName & _
                ": image Equation number is not anchored by the right tab."
    End If

    Set openingRange = documentObject.Range( _
        Start:=fieldStart - 1, _
        End:=fieldStart)
    Set closingRange = documentObject.Range( _
        Start:=fieldEnd, _
        End:=fieldEnd + 1)
    If Not nativeNumberInMath And _
       (openingRange.Text <> "(" Or closingRange.Text <> ")") Then
        Err.Raise vbObjectError + 7513, "VisualTeX", _
            assertionName & ": Equation number parentheses are incomplete."
    End If
    Set resultRange = sequenceField.Result.Duplicate
    preferredSize = resultRange.Font.Size
    If preferredSize <= 0! Or preferredSize > 72! Or _
       (Not nativeNumberInMath And _
        (openingRange.Font.Size <> preferredSize Or _
         closingRange.Font.Size <> preferredSize)) Then
        Err.Raise vbObjectError + 7514, "VisualTeX", _
            assertionName & ": Equation number font size is invalid."
    End If
    If formulaRange.OMaths.Count = 1 Then
        expectedRaise = CLng(VTNativeEquationNumberRaisePoints( _
            formulaRange, preferredSize))
    Else
        expectedRaise = CLng(VTEquationNumberRaisePoints( _
            renderedHeightPoints, preferredSize))
    End If
    If resultRange.Font.Position <> expectedRaise Or _
       (Not nativeNumberInMath And _
        (openingRange.Font.Position <> expectedRaise Or _
         closingRange.Font.Position <> expectedRaise)) Then
        Err.Raise vbObjectError + 7515, "VisualTeX", _
            assertionName & ": Equation number is not vertically centered."
    End If

    If Abs(expectedRaise) > 2 Then
        Err.Raise vbObjectError + 7515, "VisualTeX", _
            assertionName & ": native Equation number correction is too large."
    End If

    bookmarkName = VTEquationNumberBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(bookmarkName) Then
        Err.Raise vbObjectError + 7516, "VisualTeX", _
            assertionName & ": Equation number REF bookmark is missing."
    End If
    Set captionRange = documentObject.Bookmarks(bookmarkName).Range.Duplicate
    If (nativeNumberInMath And captionRange.Text <> resultRange.Text) Or _
       (Not nativeNumberInMath And _
        captionRange.Text <> "(" & resultRange.Text & ")") Or _
       InStr(1, captionRange.Text, expectedCaptionText, vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 7517, "VisualTeX", _
            assertionName & ": Equation number REF bookmark contains formula text" & _
            " [text=" & Replace$(Replace$(captionRange.Text, vbTab, "<TAB>"), vbCr, "<CR>") & _
            "; range=" & CStr(captionRange.Start) & "-" & _
            CStr(captionRange.End) & "]."
    End If

    If formulaRange.OMaths.Count = 1 Then
        If formulaRange.OMaths(1).Type <> wdOMathDisplay Or _
           formulaRange.Font.Position <> 0 Or _
           formulaRange.Font.Size <= 0! Or _
           formulaRange.Font.Size > 72! Then
            Err.Raise vbObjectError + 7517, "VisualTeX", _
                assertionName & ": numbered OMML is not native display math" & _
                " [type=" & CStr(formulaRange.OMaths(1).Type) & _
                "; size=" & CStr(formulaRange.Font.Size) & _
                "; position=" & CStr(formulaRange.Font.Position) & "]."
        End If
    End If

    documentObject.Repaginate
    Set formulaStartProbe = formulaRange.Duplicate
    formulaStartProbe.Collapse wdCollapseStart
    If nativeNumberInMath Then
        Set formulaEndProbe = documentObject.Range( _
            formulaRange.Start, equationArrayMarker.Start)
    Else
        Set formulaEndProbe = formulaRange.Duplicate
    End If
    formulaEndProbe.Collapse wdCollapseEnd
    If nativeNumberInMath Then
        Set numberEndProbe = resultRange.Duplicate
    Else
        Set numberEndProbe = closingRange.Duplicate
    End If
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
    Dim sequenceField As Field
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
    Set sequenceField = VTFindEquationSequenceField(paragraphRange)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The numbered Equation invariant field is missing."
    End If
    If paragraphRange.OMaths.Count = 1 Then
        Set formulaRange = paragraphRange.OMaths(1).Range.Duplicate
        formulaType = paragraphRange.OMaths(1).Type
    ElseIf paragraphRange.InlineShapes.Count = 1 Then
        Set formulaRange = paragraphRange.InlineShapes(1).Range.Duplicate
        formulaType = -1
    Else
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The numbered Equation invariant formula is ambiguous."
    End If

    fieldStart = VTEquationFieldStart(sequenceField)
    fieldEnd = VTEquationFieldEnd(sequenceField)
    Set openingRange = documentObject.Range(fieldStart - 1, fieldStart)
    Set closingRange = documentObject.Range(fieldEnd, fieldEnd + 1)
    Set formulaStartProbe = formulaRange.Duplicate
    formulaStartProbe.Collapse wdCollapseStart
    Set formulaEndProbe = formulaRange.Duplicate
    formulaEndProbe.Collapse wdCollapseEnd
    Set numberEndProbe = closingRange.Duplicate
    numberEndProbe.Collapse wdCollapseEnd
    bookmarkName = ""
    For Each numberBookmark In documentObject.Bookmarks
        If numberBookmark.Range.Start = openingRange.Start And _
           numberBookmark.Range.End = closingRange.End Then
            If Left$(numberBookmark.Name, _
               Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
               VT_WORD_NUMBER_BOOKMARK_PREFIX Then
                bookmarkName = numberBookmark.Name
                Exit For
            End If
        End If
    Next numberBookmark
    documentObject.Repaginate

    VTNumberedEquationInvariantSnapshot = _
        "paragraph=" & CStr(paragraphRange.Start) & ":" & _
            CStr(paragraphRange.End) & _
        "|text=" & Replace$(Replace$(paragraphRange.Text, vbTab, "<TAB>"), vbCr, "<CR>") & _
        "|field=" & CStr(fieldStart) & ":" & CStr(fieldEnd) & _
            ":" & sequenceField.Result.Text & ":" & Trim$(sequenceField.Code.Text)
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
    Set layoutTable = VTEnsureNumberedDisplayTable( _
        formulaRange, True, formulaId)
    Set formulaShape = layoutTable.Cell(1, 2).Range.InlineShapes(1)
    VTEnsureEquationNumberFields layoutTable, formulaId
    Set VTEnsureImageEquationNumber = layoutTable.Range.Duplicate
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

    If equationRange Is Nothing Or equationRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7470, "VisualTeX", _
            "The native equation number target is missing."
    End If
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    numberCreated = Not equationRange.Document.Bookmarks.Exists( _
        numberBookmarkName)
    Set formulaRange = equationRange.OMaths(1).Range.Duplicate
    Set layoutTable = VTEnsureNumberedDisplayTable( _
        formulaRange, False, formulaId)
    VTEnsureEquationNumberFields layoutTable, formulaId
    Set VTEnsureNativeEquationNumber = layoutTable.Range.Duplicate
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

    VTRefreshEquationNumberMirror _
        documentObject, sequenceField, sequenceBookmarkName, _
        sequenceOrdinal

    ' Reconcile every existing native Equation before creating this formula's
    ' visible right-cell REF. Updating SEQ fields can invalidate a Bookmark that
    ' wraps a field result on Word for Mac, so immediately re-resolve the known
    ' helper paragraph and restore the exact VT_N_/VT_C_ pair afterward.
    VTReconcileEquationNumbers documentObject
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
    ByVal documentObject As Document)

    Dim candidateField As Field
    Dim targetBookmarkName As String
    Dim formulaId As String
    Dim targetKind As String
    Dim isVisualTeXReference As Boolean

    If documentObject Is Nothing Then Exit Sub
    For Each candidateField In documentObject.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            targetBookmarkName = VTReferenceTargetBookmarkName( _
                candidateField.Code.Text)
            isVisualTeXReference = _
                (Left$(targetBookmarkName, _
                    Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)) = _
                 VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)
            If Not isVisualTeXReference And _
               Left$(targetBookmarkName, 4) = "_Ref" Then
                formulaId = VTFormulaIdForReferenceTarget( _
                    documentObject, targetBookmarkName, targetKind, _
                    candidateField.Result.Text)
                isVisualTeXReference = (Len(formulaId) > 0)
            End If
            If isVisualTeXReference Then
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

Private Sub VTReconcileEquationNumbers(ByVal documentObject As Document)
    Dim candidate As Field
    Dim candidateBookmark As Bookmark
    Dim equationLabelName As String
    Dim fieldCode As String
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim suffixText As String
    Dim sequenceOrdinal As Long

    If documentObject Is Nothing Then Exit Sub
    equationLabelName = VTNativeEquationLabelName()

    ' Phase 1: refresh every native SEQ in document order and restore the exact
    ' VT_N_ result Bookmark. The one-point white helper paragraph remains a real,
    ' non-hidden Word caption target; no second plain-text sequence is created.
    For Each candidate In documentObject.Fields
        If VTIsNativeEquationSequenceField(candidate, equationLabelName) Then
            sequenceOrdinal = sequenceOrdinal + 1
            sequenceBookmarkName = ""
            For Each candidateBookmark In documentObject.Bookmarks
                If Left$(candidateBookmark.Name, _
                   Len(VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)) = _
                   VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX Then
                    If candidateBookmark.Range.Start <= _
                       candidate.Result.Start And _
                       candidateBookmark.Range.End >= _
                       candidate.Result.End Then
                        sequenceBookmarkName = candidateBookmark.Name
                        Exit For
                    End If
                End If
            Next candidateBookmark
            If Len(sequenceBookmarkName) > 0 Then
                ' VisualTeX-owned captions migrate from the legacy restarted
                ' code to a normal flowing SEQ field. Ordinary Word Equation
                ' captions are only updated and are never rewritten.
                VTApplyEquationSequenceOrdinal _
                    candidate, equationLabelName, sequenceOrdinal
                If documentObject.Bookmarks.Exists( _
                   sequenceBookmarkName) Then
                    documentObject.Bookmarks( _
                        sequenceBookmarkName).Delete
                End If
                documentObject.Bookmarks.Add _
                    name:=sequenceBookmarkName, _
                    Range:=candidate.Result.Duplicate
                VTRefreshEquationNumberMirror _
                    documentObject, candidate, sequenceBookmarkName, _
                    sequenceOrdinal
            Else
                candidate.Update
            End If
        End If
    Next candidate

    ' Phase 2: normalize only right-cell REF fields. Each right cell must contain
    ' ordinary parentheses around a native REF to VT_N_. VT_R_ bookmarks the
    ' complete visible (n) range only after that native REF is valid.
    For Each candidateBookmark In documentObject.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_NUMBER_BOOKMARK_PREFIX Then
            If Left$(candidateBookmark.Range.Text, 1) <> "(" Or _
               Right$(candidateBookmark.Range.Text, 1) <> ")" Then
                Err.Raise vbObjectError + 7549, "VisualTeX", _
                    "A visible Equation number Bookmark changed before REF" & _
                    " normalization" & _
                    " [bookmark=" & candidateBookmark.Name & _
                    "; text=" & candidateBookmark.Range.Text & _
                    "; range=" & CStr(candidateBookmark.Range.Start) & _
                    "-" & CStr(candidateBookmark.Range.End) & "]."
            End If
        End If
    Next candidateBookmark

    For Each candidate In documentObject.Fields
        If candidate.Type = wdFieldRef Then
            If candidate.Result.Information(wdWithInTable) Then
                fieldCode = candidate.Code.Text
                sequenceBookmarkName = VTBookmarkNameInFieldCode( _
                    documentObject, fieldCode, _
                    VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX)
                If Len(sequenceBookmarkName) > 0 Then
                    VTNormalizeVisibleEquationReferenceField _
                        documentObject, candidate, sequenceBookmarkName
                End If
            End If
        End If
    Next candidate

    For Each candidateBookmark In documentObject.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_NUMBER_BOOKMARK_PREFIX Then
            If Left$(candidateBookmark.Range.Text, 1) <> "(" Or _
               Right$(candidateBookmark.Range.Text, 1) <> ")" Then
                Err.Raise vbObjectError + 7549, "VisualTeX", _
                    "A visible Equation number Bookmark changed during REF" & _
                    " normalization" & _
                    " [bookmark=" & candidateBookmark.Name & _
                    "; text=" & candidateBookmark.Range.Text & _
                    "; range=" & CStr(candidateBookmark.Range.Start) & _
                    "-" & CStr(candidateBookmark.Range.End) & "]."
            End If
        End If
    Next candidateBookmark

    ' Phase 3: refresh every native body and right-cell REF. Word inherits the
    ' one-point SEQ helper formatting when a native Equation cross-reference is
    ' inserted or updated, so VisualTeX-owned body REF results must be restored
    ' to the document's visible body size after every field refresh.
    For Each candidate In documentObject.Fields
        If candidate.Type = wdFieldRef Then
            candidate.Update
            With candidate.Result.Font
                .Hidden = False
                .Color = wdColorAutomatic
                .Position = 0
                If candidate.Result.Information(wdWithInTable) Then
                    .Size = VTVisibleEquationNumberFontSize(documentObject)
                End If
            End With
        End If
    Next candidate
    VTNormalizeBodyEquationReferenceVisibility documentObject

    For Each candidateBookmark In documentObject.Bookmarks
        If Left$(candidateBookmark.Name, _
           Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_NUMBER_BOOKMARK_PREFIX Then
            If Left$(candidateBookmark.Range.Text, 1) <> "(" Or _
               Right$(candidateBookmark.Range.Text, 1) <> ")" Then
                Err.Raise vbObjectError + 7549, "VisualTeX", _
                    "A visible Equation number Bookmark changed while refreshing REF" & _
                    " fields" & _
                    " [bookmark=" & candidateBookmark.Name & _
                    "; text=" & candidateBookmark.Range.Text & _
                    "; range=" & CStr(candidateBookmark.Range.Start) & _
                    "-" & CStr(candidateBookmark.Range.End) & "]."
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
                "Word lost the numbered formula table after native conversion."
        End If
        Set numberLayoutRange = targetDocument.Bookmarks( _
            VTEquationNumberBookmarkName(formulaId)).Range.Tables(1). _
            Cell(1, 2).Range.Duplicate
        If numberLayoutRange.OMaths.Count <> 1 Then
            Err.Raise vbObjectError + 7460, "VisualTeX", _
                "Word lost the converted OMath while refreshing its number."
        End If
        Set finalFormulaRange = _
            numberLayoutRange.OMaths(1).Range.Duplicate
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
        If Not finalFormulaRange.Information(wdWithInTable) Or _
           finalFormulaRange.Tables.Count <> 1 Then
            Err.Raise vbObjectError + 7460, "VisualTeX", _
                "The converted OMath left its numbered table before finalization."
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
