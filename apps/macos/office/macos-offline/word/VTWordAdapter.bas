Attribute VB_Name = "VTWordAdapter"
Option Explicit

#Const VT_WORD_DIAGNOSTICS = False

' The macros Word, the Ribbon, the app and AppleScript call (every Public procedure of the Word add-in's
' VBA modules is here or in VTRibbonCallbacks). Their code is in the VTWord* modules, which are private to
' the project (Option Private Module).

Private Const VT_WORD_IMAGE_FONT_CONTROL_ID As String = _
    "VisualTeX.Mac.Word.ImageFontSize"
Private Const VT_WORD_NUMBERING_FORMAT_CONTROL_ID As String = _
    "VisualTeX.Mac.Word.NumberingFormat"
Private Const VT_WORD_DOUBLE_CLICK_TRACE_FILE As String = _
    "/word-double-click-trace.log"
Private VT_WORD_EVENT_SINK As VTWordEvents
Private VT_WORD_APPLICATION_PREWARM_SCHEDULED As Boolean
Private VT_WORD_APPLICATION_PREWARM_RUNNING As Boolean
Private VT_WORD_APPLICATION_PREWARM_ATTEMPTED As Boolean
Private VT_WORD_DOUBLE_CLICK_TRACE_CONTEXT_DEPTH As Long
Private VT_WORD_DOUBLE_CLICK_TRACE_BUFFER As String
Private VT_WORD_DOUBLE_CLICK_TRACE_STARTED_AT As Single
Private VT_WORD_IMAGE_MACRO_MIGRATION_SCHEDULED As Boolean
Private VT_WORD_IMAGE_MACRO_MIGRATION_RUNNING As Boolean
Private VT_WORD_IMAGE_WATCH_SECONDS_SINCE_SCAN As Long
Private VT_WORD_IDLE_TASKS_RUNNING As Boolean
Private VT_WORD_TYPOGRAPHY_FINGERPRINT As String
Private VT_WORD_DOCUMENT_TYPOGRAPHY_FINGERPRINT As String
Private VT_WORD_SELECTED_IMAGE_FINGERPRINT As String
Private VT_WORD_IMAGE_FONT_CONTROL_STATE As String
Private VT_WORD_REFERENCE_MENU_DOCUMENT As Document
Private VT_WORD_REFERENCE_MENU_INSERTION_RANGE As Range

' AI agents: the entry macro the app runs (see VTWordAgent)
Public Sub VisualTeX_AgentRun()
    Dim requestId As String
    Dim operation As String
    Dim data As String
    Dim code As String
    Dim message As String
    Dim separator As Long

    On Error GoTo AgentFailed
    VT_AGENT_ERRORS = ""
    VT_AGENT_NOTES = ""
    VT_AGENT_SESSION_ID = ""
    VT_AGENT_REQUEST_JSON = ""
    Set VT_AGENT_REQUEST = VTAgentReadRequest()
    requestId = VTAgentRaw("id")
    operation = VTAgentRaw("op")
    VT_AGENT_MODE = True
    Select Case operation
        Case "list_documents": data = VTAgentListDocuments()
        Case "get_selection": data = VTAgentGetSelection()
        Case "list_formulas": data = VTAgentListFormulas()
        Case "read_formula": data = VTAgentReadFormula()
        Case "find_vtx_formula": data = VTAgentFindVisualTeXFormula()
        Case "begin_insert": data = VTAgentBeginInsert()
        Case "begin_update": data = VTAgentBeginUpdate()
        Case "begin_import": data = VTAgentBeginImport()
        Case "select_formula": data = VTAgentSelectFormulaOperation()
        Case "delete_formula": data = VTAgentDeleteFormula()
        Case "set_formula_font_size": data = VTAgentSetFontSize()
        Case "number_display_formulas": data = VTAgentNumberDisplayFormulas()
        Case "update_equation_numbers": data = VTAgentUpdateEquationNumbers()
        Case "set_equation_number_format": data = VTAgentSetNumberFormat()
        Case "insert_equation_reference": data = VTAgentInsertReference()
        Case "latex_text_to_formulas", "formulas_to_latex_text", "convert_formulas"
            data = VTAgentDocumentOperation(operation)
        Case Else
            Err.Raise vbObjectError + 7900, "VisualTeX", "unknown_operation|Word has no agent operation " & operation & "."
    End Select
    VT_AGENT_MODE = False
    VTAgentWriteResponse "{""id"":" & VTJsonString(requestId) & ",""ok"":true,""data"":" & data & "}"
    Exit Sub

AgentFailed:
    message = Err.Description
    VT_AGENT_MODE = False
    code = "operation_failed"
    separator = InStr(1, message, "|", vbBinaryCompare)
    If separator > 1 And separator < 40 And InStr(1, Left$(message, separator - 1), " ", vbBinaryCompare) = 0 Then
        code = Left$(message, separator - 1)
        message = Mid$(message, separator + 1)
    End If
    If Len(VT_AGENT_ERRORS) > 0 Then message = message & " " & VT_AGENT_ERRORS
    On Error Resume Next
    VTAgentWriteResponse "{""id"":" & VTJsonString(requestId) & ",""ok"":false,""error"":{""code"":" & _
        VTJsonString(code) & ",""message"":" & VTJsonString(message) & "}}"
End Sub

Public Sub VisualTeX_InitializeWordHost()
    If VTWordVbeBuildTemplateActive() Then Exit Sub
    On Error Resume Next
    VTInitializeWordDoubleClickTrace
    VTInitializeWordEvents
    ' Keep host initialization explicit instead of exposing Word's special
    ' AutoExec startup macro. On Word for Mac, a Startup-template AutoExec can
    ' run before Finder's first open-document AppleEvent is consumed, leaving
    ' Word open with no document until the user double-clicks the file again.
    VTEnsureApplicationPrewarmScheduled
    VisualTeX_MigrateImageMacroButtons
    VTRepairOpenedWpsDocuments
    VTEnsureOrphanWatchScheduled
    If Documents.Count > 0 Then _
        VTSynchronizeWordImageTypography ActiveDocument.Content
    VTEnsureImageSizeWatchScheduled
    VTWriteWordHealth
    On Error GoTo 0
End Sub

Public Sub VTEnsureApplicationPrewarmScheduled()
    If VTWordVbeBuildTemplateActive() Then Exit Sub
    On Error GoTo ScheduleFailed
    If Documents.Count = 0 Then Exit Sub
    If VT_WORD_APPLICATION_PREWARM_SCHEDULED Or _
       VT_WORD_APPLICATION_PREWARM_RUNNING Or _
       VT_WORD_APPLICATION_PREWARM_ATTEMPTED Then Exit Sub

    VT_WORD_APPLICATION_PREWARM_SCHEDULED = True
    VTScheduleWordIdleTasks 1
    Exit Sub

ScheduleFailed:
    VT_WORD_APPLICATION_PREWARM_SCHEDULED = False
End Sub

Public Sub VisualTeX_PrewarmWordApplication()
    If VTWordVbeBuildTemplateActive() Then Exit Sub
    VT_WORD_APPLICATION_PREWARM_SCHEDULED = False
    If VT_WORD_APPLICATION_PREWARM_RUNNING Or _
       VT_WORD_APPLICATION_PREWARM_ATTEMPTED Then Exit Sub

    VT_WORD_APPLICATION_PREWARM_RUNNING = True
    VT_WORD_APPLICATION_PREWARM_ATTEMPTED = True
    On Error Resume Next
    VTPrewarmApplication VT_WORD_HOST
    On Error GoTo 0
    VT_WORD_APPLICATION_PREWARM_RUNNING = False
End Sub

Public Sub AutoOpen()
    If VTWordVbeBuildTemplateActive() Then Exit Sub
    On Error Resume Next
    ' AutoOpen runs after the first document exists. Unlike AutoExec, it cannot
    ' consume Finder's initial document-open request before Word creates it.
    VisualTeX_InitializeWordHost
    On Error GoTo 0
End Sub

Public Function VTWordSourceSelfTest() As Boolean
    Dim equationLabelName As String
    Dim optionalDispatch As Collection

    If VTBase64UrlDecodeUtf8("XGZyYWN7YX17Yn0") <> "\frac{a}{b}" Then Exit Function
    If VTLaTeXToWordLinear("\frac{a_1}{\sqrt{x}}") <> "(a_(1))/(\sqrt(x))" Then Exit Function
    If VTLaTeXToWordLinear("\alpha+\beta") <> "\alpha+\beta" Then Exit Function
    equationLabelName = VTNativeEquationLabelName()
    If Len(equationLabelName) = 0 Then Exit Function
    If InStr(1, VTEquationSequenceFieldText(equationLabelName), equationLabelName, vbTextCompare) = 0 Then Exit Function
    ' Older renderers omit optional geometry. Missing and empty values must
    ' retain their defaults without indexing a missing Collection member.
    Set optionalDispatch = New Collection
    If VTDispatchOptionalDouble(optionalDispatch, "inkCenterYRatio", 0.5) <> 0.5 Then Exit Function
    optionalDispatch.Add "", "inkCenterYRatio"
    If VTDispatchOptionalDouble(optionalDispatch, "inkCenterYRatio", 0.5) <> 0.5 Then Exit Function
    optionalDispatch.Remove "inkCenterYRatio"
    optionalDispatch.Add "0.47826087", "inkCenterYRatio"
    If Abs(VTDispatchOptionalDouble(optionalDispatch, "inkCenterYRatio", 0.5) - 0.47826087) > 0.00000001 Then Exit Function
    VTWordSourceSelfTest = True
End Function

Public Sub VisualTeX_PerformanceNoop()
End Sub

Public Function VTWordInternalMutationActive() As Boolean
    VTWordInternalMutationActive = _
        (VT_WORD_INTERNAL_MUTATION_DEPTH > 0)
End Function

Public Sub VTInitializeWordEvents()
    On Error GoTo InitializationFailed
    If Not VT_WORD_EVENT_SINK Is Nothing Then Exit Sub
    VTTraceWordDoubleClick "events-initialize-enter", Nothing, ""
    Set VT_WORD_EVENT_SINK = New VTWordEvents
    Set VT_WORD_EVENT_SINK.App = Word.Application
    VTTraceWordDoubleClick _
        "events-initialize-complete", Nothing, _
        "sinkReady=" & CStr(Not VT_WORD_EVENT_SINK Is Nothing)
    Exit Sub

InitializationFailed:
    VTTraceWordDoubleClick _
        "events-initialize-error", Nothing, _
        "error=" & CStr(Err.Number) & ":" & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub VisualTeX_StabilizeImageEquationNumberSelection( _
    ByVal selectedRange As Range)

    Dim documentObject As Document
    Dim paragraphRange As Range
    Dim formulaRange As Range
    Dim visibleNumberField As Field
    Dim sequenceField As Field
    Dim candidateField As Field
    Dim targetBookmarkName As String
    Dim formulaId As String
    Dim mutationStarted As Boolean

    On Error GoTo StabilizeFailed
    If selectedRange Is Nothing Or VTWordInternalMutationActive() Then Exit Sub
    Set documentObject = selectedRange.Document
    Set paragraphRange = selectedRange.Paragraphs(1).Range.Duplicate

    For Each candidateField In paragraphRange.Fields
        If candidateField.Type = wdFieldRef Then
            If candidateField.Result.Start <= selectedRange.End And _
               candidateField.Result.End >= selectedRange.Start Then
                targetBookmarkName = VTReferenceTargetBookmarkName( _
                    candidateField.Code.Text)
                formulaId = VTFormulaIdFromSequenceBookmarkName( _
                    targetBookmarkName)
                If Len(formulaId) > 0 Then
                    Set formulaRange = VTNumberedFormulaRangeForId( _
                        documentObject, formulaId)
                    If Not formulaRange Is Nothing Then
                        If formulaRange.InlineShapes.Count = 1 And _
                           formulaRange.OMaths.Count = 0 Then
                            Set visibleNumberField = _
                                VTImageEquationReferenceField( _
                                    formulaRange, formulaId)
                            If Not visibleNumberField Is Nothing Then
                                If VTEquationFieldStart(visibleNumberField) = _
                                   VTEquationFieldStart(candidateField) Then
                                    Set sequenceField = _
                                        VTNativeEquationSequenceHelperField( _
                                            documentObject, formulaId)
                                    If Not sequenceField Is Nothing Then
                                        VTBeginWordInternalMutation
                                        mutationStarted = True
                                        VT_WORD_KEEP_NUMBER_FORMAT = True
                                        VTRefreshParagraphEquationBookmarks _
                                            documentObject, sequenceField, formulaId
                                        VT_WORD_KEEP_NUMBER_FORMAT = False
                                        VTEndWordInternalMutation
                                        mutationStarted = False
                                    End If
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next candidateField
    Exit Sub

StabilizeFailed:
    On Error Resume Next
    VT_WORD_KEEP_NUMBER_FORMAT = False
    If mutationStarted Then VTEndWordInternalMutation
    On Error GoTo 0
End Sub

Public Function VTRepairSelectedWordNativeEquationCrossReference( _
    ByVal selectedRange As Range, _
    Optional ByRef diagnosticText As String) As Boolean

    Dim documentObject As Document
    Dim selectedParagraph As Range
    Dim candidateField As Field
    Dim targetRange As Range
    Dim sequenceRange As Range
    Dim formulaIds As Variant
    Dim formulaId As String
    Dim targetBookmarkName As String
    Dim targetKind As String
    Dim sequenceBookmarkName As String
    Dim expectedNumber As String
    Dim repairedCount As Long
    Dim internalMutationStarted As Boolean
    Dim repairErrorNumber As Long
    Dim repairErrorDescription As String

    diagnosticText = "entry"
    If selectedRange Is Nothing Then
        diagnosticText = diagnosticText & "; noRange"
        Exit Function
    End If
    If VTWordInternalMutationActive() Then
        diagnosticText = diagnosticText & "; mutationActive"
        Exit Function
    End If

    Set documentObject = selectedRange.Document
    Set selectedParagraph = selectedRange.Paragraphs(1).Range.Duplicate
    diagnosticText = diagnosticText & _
        "; selected=" & CStr(selectedRange.Start) & "-" & _
            CStr(selectedRange.End) & _
        "; paragraph=" & CStr(selectedParagraph.Start) & "-" & _
            CStr(selectedParagraph.End) & _
        "; fields=" & CStr(selectedParagraph.Fields.Count)
    If selectedParagraph.Information(wdWithInTable) Then
        diagnosticText = diagnosticText & "; inTable"
        Exit Function
    End If

    On Error GoTo RepairFailed
    For Each candidateField In selectedParagraph.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            targetBookmarkName = VTReferenceTargetBookmarkName( _
                candidateField.Code.Text)
            diagnosticText = diagnosticText & _
                "; ref=" & targetBookmarkName & _
                "/result[" & Trim$(candidateField.Result.Text) & "]"
            If Left$(targetBookmarkName, 4) = "_Ref" Then
                ' SelectionChange also runs for ordinary typing and formula
                ' edits. Enumerate document-wide identities only when this
                ' local paragraph actually contains a native private REF.
                If Not IsArray(formulaIds) Then
                    formulaIds = VTValidNumberedFormulaIds(documentObject)
                    diagnosticText = diagnosticText & _
                        "; formulas=" & CStr(VTVariantArrayCount(formulaIds))
                    If Not IsArray(formulaIds) Then Exit Function
                End If
                Set targetRange = Nothing
                If VTTryGetBookmarkRangeIncludingHidden( _
                   documentObject, targetBookmarkName, targetRange) Then
                    diagnosticText = diagnosticText & _
                        "/target=" & CStr(targetRange.Start) & "-" & _
                            CStr(targetRange.End)
                    targetKind = ""
                    formulaId = VTFormulaIdForReferenceTarget( _
                        documentObject, targetBookmarkName, targetKind, _
                        candidateField.Result.Text, formulaIds)
                    diagnosticText = diagnosticText & _
                        "/formula=" & formulaId & "/kind=" & targetKind
                    If Len(formulaId) > 0 Then
                        sequenceBookmarkName = _
                            VTEquationSequenceNumberBookmarkName(formulaId)
                        If Not documentObject.Bookmarks.Exists( _
                           sequenceBookmarkName) Then
                            Err.Raise vbObjectError + 7555, "VisualTeX", _
                                "The native Equation cross-reference target lost " & _
                                "its external SEQ helper."
                        End If
                        Set sequenceRange = documentObject.Bookmarks( _
                            sequenceBookmarkName).Range.Duplicate
                        expectedNumber = VTEquationNumberTextForFormula( _
                            documentObject, formulaId)
                        diagnosticText = diagnosticText & _
                            "/sequence=" & CStr(sequenceRange.Start) & "-" & _
                                CStr(sequenceRange.End) & _
                            "/expected[" & expectedNumber & "]"
                        If Len(expectedNumber) = 0 Then
                            Err.Raise vbObjectError + 7555, "VisualTeX", _
                                "The native Equation cross-reference has no " & _
                                "formatted number."
                        End If

                        ' Selection changes in an already-correct reference
                        ' must not update the field and reformat its result.
                        If targetRange.Start = sequenceRange.Start And _
                           targetRange.End = sequenceRange.End And _
                           VTReferenceResultMatchesEquationNumber( _
                               candidateField.Result.Text, expectedNumber) And _
                           VTBodyEquationReferenceFormattingMatches( _
                               documentObject, candidateField.Result) Then _
                            GoTo ReferenceAlreadyCorrect

                        If Not internalMutationStarted Then
                            VTBeginWordInternalMutation
                            internalMutationStarted = True
                        End If
                        If targetRange.Start <> sequenceRange.Start Or _
                           targetRange.End <> sequenceRange.End Then
                            VTReplaceBookmarkRangeIncludingHidden _
                                documentObject, targetBookmarkName, sequenceRange
                        End If
                        candidateField.Update
                        diagnosticText = diagnosticText & _
                            "/updated[" & Trim$(candidateField.Result.Text) & "]"
                        VTFormatBodyEquationReference _
                            documentObject, candidateField, _
                            candidateField.Result.Duplicate
                        If Not VTReferenceResultMatchesEquationNumber( _
                           candidateField.Result.Text, expectedNumber) Then
                            Err.Raise vbObjectError + 7555, "VisualTeX", _
                                "Word did not repair its native Equation " & _
                                "cross-reference [target=" & targetBookmarkName & _
                                "; result=" & candidateField.Result.Text & _
                                "; expected=" & expectedNumber & "]."
                        End If
                        repairedCount = repairedCount + 1
ReferenceAlreadyCorrect:
                    End If
                Else
                    diagnosticText = diagnosticText & "/bookmarkMissing"
                End If
            End If
        End If
    Next candidateField

    If internalMutationStarted Then VTEndWordInternalMutation
    diagnosticText = diagnosticText & _
        "; repaired=" & CStr(repairedCount)
    VTRepairSelectedWordNativeEquationCrossReference = _
        (repairedCount > 0)
    Exit Function

RepairFailed:
    repairErrorNumber = Err.Number
    repairErrorDescription = Err.Description
    diagnosticText = diagnosticText & _
        "; error=" & CStr(repairErrorNumber) & ":" & repairErrorDescription
    On Error Resume Next
    If internalMutationStarted Then VTEndWordInternalMutation
    On Error GoTo 0
    Err.Raise repairErrorNumber, _
        "VisualTeX native Equation cross-reference repair", _
        repairErrorDescription
End Function

Public Function VTRepairNumberedImageReturnAfterFormula( _
    ByVal selectedRange As Range) As Boolean

    Dim documentObject As Document
    Dim blankParagraph As Range
    Dim previousProbe As Range
    Dim formulaParagraph As Range
    Dim formulaRange As Range
    Dim formulaShape As InlineShape
    Dim sequenceField As Field
    Dim helperParagraph As Range
    Dim typingParagraph As Range
    Dim insertionRange As Range
    Dim visibleNumberField As Field
    Dim numberRange As Range
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim originalBlankStart As Long
    Dim blankDeleted As Boolean
    Dim typingCreated As Boolean
    Dim internalMutationStarted As Boolean
    Dim repairErrorNumber As Long

    If selectedRange Is Nothing Or _
       VTWordInternalMutationActive() Then Exit Function
    If selectedRange.Start <> selectedRange.End Then Exit Function
    Set documentObject = selectedRange.Document
    Set blankParagraph = selectedRange.Paragraphs(1).Range.Duplicate
    If blankParagraph.Information(wdWithInTable) Or _
       blankParagraph.Fields.Count <> 0 Or _
       blankParagraph.InlineShapes.Count <> 0 Or _
       blankParagraph.OMaths.Count <> 0 Or _
       VTWordRangeHasMeaningfulText(blankParagraph) Then Exit Function
    If blankParagraph.Start <= 0 Then Exit Function

    ' Word inserts Return immediately after the visible formula row, which pushes
    ' the external SEQ helper one paragraph away. Match only that exact transient
    ' structure: formula -> one empty paragraph -> this formula's helper.
    Set previousProbe = documentObject.Range( _
        Start:=blankParagraph.Start - 1, End:=blankParagraph.Start)
    Set formulaParagraph = previousProbe.Paragraphs(1).Range.Duplicate
    If formulaParagraph.End <> blankParagraph.Start Or _
       formulaParagraph.Information(wdWithInTable) Or _
       formulaParagraph.InlineShapes.Count <> 1 Or _
       formulaParagraph.OMaths.Count <> 0 Then Exit Function
    Set formulaShape = formulaParagraph.InlineShapes(1)
    If Not VTTryParseFormulaReference( _
       formulaShape.Title, formulaId, displayMode, numbered) Then Exit Function
    If displayMode <> "block" Or Not numbered Then Exit Function
    Set formulaRange = formulaShape.Range.Duplicate
    Set visibleNumberField = VTImageEquationReferenceField( _
        formulaRange, formulaId)
    If visibleNumberField Is Nothing Then Exit Function
    Set numberRange = VTImageEquationNumberRange( _
        formulaRange, visibleNumberField)
    If numberRange Is Nothing Then Exit Function

    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    If sequenceField Is Nothing Then Exit Function
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If Not VTHelperParagraphOwnsNativeEquationSequence( _
           helperParagraph) Or _
       helperParagraph.Start <> blankParagraph.End Then Exit Function

    originalBlankStart = blankParagraph.Start
    On Error GoTo RepairFailed
    VTBeginWordInternalMutation
    internalMutationStarted = True

    ' Remove only the just-created empty line, restoring formula/helper adjacency.
    blankParagraph.Delete
    blankDeleted = True
    Set formulaRange = VTNumberedFormulaRangeForId( _
        documentObject, formulaId)
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    If formulaRange Is Nothing Or sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word lost the numbered image while repairing Return."
    End If
    Set formulaParagraph = _
        VTWordParagraphContainingFormula(formulaRange)
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    If formulaParagraph Is Nothing Or _
       helperParagraph.Start <> formulaParagraph.End Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not restore image/helper adjacency after Return."
    End If

    ' Recreate the user's requested line after the helper. Word can inherit the
    ' helper's compact formatting/Frame-like paragraph state, so normalize the
    ' new line explicitly to ordinary body typing before placing the caret there.
    helperParagraph.InsertParagraphAfter
    Set sequenceField = VTNativeEquationSequenceHelperField( _
        documentObject, formulaId)
    If sequenceField Is Nothing Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word lost the image Equation SEQ while creating the typing line."
    End If
    Set helperParagraph = _
        sequenceField.Result.Paragraphs(1).Range.Duplicate
    VTSetCollapsedEquationCaptionBookmark _
        documentObject, formulaId, helperParagraph
    VTFormatHiddenEquationParagraph helperParagraph
    If helperParagraph.End >= documentObject.Content.End Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word did not create a paragraph after the image Equation helper."
    End If
    Set insertionRange = documentObject.Range( _
        Start:=helperParagraph.End, End:=helperParagraph.End)
    Set typingParagraph = insertionRange.Paragraphs(1).Range.Duplicate
    If typingParagraph.Start <> helperParagraph.End Or _
       typingParagraph.Information(wdWithInTable) Then
        Err.Raise vbObjectError + 7598, "VisualTeX", _
            "Word created the image Equation typing line at an invalid boundary."
    End If
    VTNormalizePlainWordParagraph typingParagraph
    typingCreated = True
    Set insertionRange = documentObject.Range( _
        Start:=typingParagraph.Start, End:=typingParagraph.Start)
    insertionRange.Select

    VTEndWordInternalMutation
    internalMutationStarted = False
    VTRepairNumberedImageReturnAfterFormula = True
    Exit Function

RepairFailed:
    repairErrorNumber = Err.Number
    On Error Resume Next
    If blankDeleted And Not typingCreated Then
        ' Preserve ordinary Word behavior if the narrow repair cannot complete:
        ' restore one empty paragraph at the original pre-helper boundary.
        Set sequenceField = VTNativeEquationSequenceHelperField( _
            documentObject, formulaId)
        If Not sequenceField Is Nothing Then
            Set helperParagraph = _
                sequenceField.Result.Paragraphs(1).Range.Duplicate
            If helperParagraph.Start = originalBlankStart Then
                Set insertionRange = documentObject.Range( _
                    Start:=helperParagraph.Start, End:=helperParagraph.Start)
                insertionRange.InsertParagraphBefore
            End If
        End If
    End If
    If internalMutationStarted Then VTEndWordInternalMutation
    Err.Clear
    On Error GoTo 0
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

Public Sub VisualTeX_RestoreSelectionOmmlToLatex()
    Dim targetRange As Range

    On Error GoTo Failed
    VTRequireWritableWordDocument
    If Selection.Range.Start = Selection.Range.End Then
        Set targetRange = VTFormulaRestoreSelectionRange("omml")
    Else
        Set targetRange = Selection.Range.Duplicate
    End If
    If VTRestoreCachedNativeFormulasToLatex(targetRange) Then Exit Sub
    VTStartWordFormulaRestore "selection", "omml", "latex"
    Exit Sub

Failed:
    VTShowError "Word OMML-to-LaTeX restore", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_RestoreSelectionImageToLatex()
    Dim targetRange As Range

    On Error GoTo Failed
    VTRequireWritableWordDocument
    If Selection.Range.Start = Selection.Range.End Then
        Set targetRange = VTFormulaRestoreSelectionRange("image")
    Else
        Set targetRange = Selection.Range.Duplicate
    End If
    If VTRestoreCachedImageFormulasToLatex(targetRange) Then Exit Sub
    VTStartWordFormulaRestore "selection", "image", "latex"
    Exit Sub

Failed:
    VTShowError "Word image-to-LaTeX restore", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_RestoreDocumentOmmlToLatex()
    On Error GoTo Failed
    VTRequireWritableWordDocument
    If VTRestoreCachedNativeFormulasToLatex( _
       ActiveDocument.Content.Duplicate) Then Exit Sub
    VTStartWordFormulaRestore "document", "omml", "latex"
    Exit Sub

Failed:
    VTShowError "Word OMML-to-LaTeX restore", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_RestoreDocumentImageToLatex()
    On Error GoTo Failed
    VTRequireWritableWordDocument
    If VTRestoreCachedImageFormulasToLatex( _
       ActiveDocument.Content.Duplicate) Then Exit Sub
    VTStartWordFormulaRestore "document", "image", "latex"
    Exit Sub

Failed:
    VTShowError "Word image-to-LaTeX restore", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_ConvertDocumentImagesToOmml()
    VTStartWordFormulaRestore "document", "image", "omml"
End Sub

Public Sub VisualTeX_ConvertDocumentOmmlToImages()
    VTStartWordFormulaRestore "document", "omml", "image"
End Sub

Public Sub VisualTeX_RedrawSelectionToImage()
    VTStartWordLatexRedraw "selection", "image"
End Sub

Public Sub VisualTeX_RedrawSelectionToOmml()
    VTStartWordLatexRedraw "selection", "omml"
End Sub

Public Sub VisualTeX_RedrawDocumentToImage()
    VTStartWordLatexRedraw "document", "image"
End Sub

Public Sub VisualTeX_RedrawDocumentToOmml()
    VTStartWordLatexRedraw "document", "omml"
End Sub

Public Sub VisualTeX_InsertLatexMarkdownDocument()
    Dim sessionId As String
    Dim bookmarkName As String
    Dim sourceDocumentId As String
    Dim insertionRange As Range
    Dim defaultFontSizePt As Double
    Dim requestJson As String
    Dim launchTiming As String
    Dim operationStage As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed
    operationStage = "require-document"
    VTRequireWritableWordDocument
    operationStage = "refresh-health"
    VTRefreshWordHealthQuietly
    operationStage = "cleanup-selection"
    VTCleanupOrphanedNumberedDisplaySelection Selection.Range

    operationStage = "create-identities"
    sessionId = VTNewUuidV4()
    bookmarkName = VTDocumentImportBookmarkName(sessionId)
    sourceDocumentId = VTWordDocumentIdentity()
    defaultFontSizePt = VTV2CurrentTypingFontSize()
    Set insertionRange = Selection.Range.Duplicate
    insertionRange.Collapse wdCollapseStart
    If ActiveDocument.Bookmarks.Exists(bookmarkName) Then
        ActiveDocument.Bookmarks(bookmarkName).Delete
    End If
    ActiveDocument.Bookmarks.Add Name:=bookmarkName, Range:=insertionRange

    operationStage = "serialize-request"
    requestJson = VTDocumentImportRequestJson( _
        sessionId, sourceDocumentId, bookmarkName, defaultFontSizePt)
    requestJson = VTWithConversionFontSizes(requestJson)
    requestJson = VTV2WithDocumentMathFont(requestJson, ActiveDocument)
    operationStage = "write-and-launch"
    launchTiming = _
        VTWriteAndLaunchSession(VT_WORD_HOST, sessionId, requestJson)
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
        "Word LaTeX/Markdown import (" & operationStage & ")", _
        errorNumber, errorDescription
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

Public Sub VisualTeX_CreateNativeNumberedDisplay()
    VTWordCreate "block", True, True
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
    Set VT_WORD_RIBBON = ribbon
    VT_WORD_IMAGE_FONT_CONTROL_STATE = ""
    VTInitializeWordDoubleClickTrace
    VTInitializeWordEvents
    VTEnsureApplicationPrewarmScheduled
    VisualTeX_MigrateImageMacroButtons
    VTRepairOpenedWpsDocuments
    VTEnsureOrphanWatchScheduled
    VTEnsureImageSizeWatchScheduled
    VTInvalidateWordImageFontSizeControl
    VTInvalidateWordEquationNumberingFormatControl
    VTRefreshWordHealthQuietly
End Sub

Public Sub VTWordRibbonDocumentImport(ByVal control As IRibbonControl)
    VisualTeX_InsertLatexMarkdownDocument
End Sub

Public Sub VTWordRibbonRedrawSelectionImage( _
    ByVal control As IRibbonControl)
    VisualTeX_RedrawSelectionToImage
End Sub

Public Sub VTWordRibbonRedrawSelectionOmml( _
    ByVal control As IRibbonControl)
    VisualTeX_RedrawSelectionToOmml
End Sub

Public Sub VTWordRibbonRedrawDocumentImage( _
    ByVal control As IRibbonControl)
    VisualTeX_RedrawDocumentToImage
End Sub

Public Sub VTWordRibbonRedrawDocumentOmml( _
    ByVal control As IRibbonControl)
    VisualTeX_RedrawDocumentToOmml
End Sub

Public Sub VTWordRibbonRestoreSelectionOmmlToLatex( _
    ByVal control As IRibbonControl)
    VisualTeX_RestoreSelectionOmmlToLatex
End Sub

Public Sub VTWordRibbonRestoreSelectionImageToLatex( _
    ByVal control As IRibbonControl)
    VisualTeX_RestoreSelectionImageToLatex
End Sub

Public Sub VTWordRibbonRestoreDocumentOmmlToLatex( _
    ByVal control As IRibbonControl)
    VisualTeX_RestoreDocumentOmmlToLatex
End Sub

Public Sub VTWordRibbonRestoreDocumentImageToLatex( _
    ByVal control As IRibbonControl)
    VisualTeX_RestoreDocumentImageToLatex
End Sub

Public Sub VTWordRibbonConvertDocumentImagesToOmml( _
    ByVal control As IRibbonControl)
    VisualTeX_ConvertDocumentImagesToOmml
End Sub

Public Sub VTWordRibbonConvertDocumentOmmlToImages( _
    ByVal control As IRibbonControl)
    VisualTeX_ConvertDocumentOmmlToImages
End Sub

Public Sub VTWordRibbonInline(ByVal control As IRibbonControl)
    VisualTeX_CreateInline
End Sub

Public Sub VTWordRibbonDisplay(ByVal control As IRibbonControl)
    VisualTeX_CreateDisplay
End Sub

Public Sub VTWordRibbonEdit(ByVal control As IRibbonControl)
    VisualTeX_EditSelected
End Sub

Public Sub VTWordRibbonConvertNative(ByVal control As IRibbonControl)
    VisualTeX_ConvertSelectedToNativeEquation
End Sub

Public Sub VTWordRibbonNumbering(ByVal control As IRibbonControl)
    VisualTeX_UpdateEquationNumbers
End Sub

Public Sub VTWordRibbonUpdateEquationFormat(ByVal control As IRibbonControl)
    VisualTeX_UpdateEquationFormat
End Sub

Public Sub VTWordRibbonNumberDisplayFormulasSelection( _
    ByVal control As IRibbonControl)
    VisualTeX_NumberDisplayFormulasSelection
End Sub

Public Sub VTWordRibbonNumberDisplayFormulasDocument( _
    ByVal control As IRibbonControl)
    VisualTeX_NumberDisplayFormulasDocument
End Sub

Public Sub VTWordRibbonGetNumberingFormatItemCount( _
    ByVal control As IRibbonControl, _
    ByRef returnedValue)

    returnedValue = 5
End Sub

Public Sub VTWordRibbonGetNumberingFormatItemLabel( _
    ByVal control As IRibbonControl, _
    ByVal itemIndex As Integer, _
    ByRef returnedValue)

    Select Case itemIndex
        Case 0
            returnedValue = VTUnicodeText(39034, 24207, 32534, 21495) & _
                " (1)"
        Case 1
            returnedValue = VTUnicodeText(25353, 31456, 32534, 21495) & _
                " (2.1)"
        Case 2
            returnedValue = VTUnicodeText(25353, 31456, 32534, 21495) & _
                " (2" & ChrW(8208) & "1)"
        Case 3
            returnedValue = VTUnicodeText(25353, 33410, 32534, 21495) & _
                " (2.3.1)"
        Case 4
            returnedValue = VTUnicodeText(25353, 33410, 32534, 21495) & _
                " (2.3" & ChrW(8208) & "1)"
        Case Else
            returnedValue = VTUnicodeText(32534, 21495, 26684, 24335)
    End Select
End Sub

Public Sub VTWordRibbonGetNumberingFormatSelectedIndex( _
    ByVal control As IRibbonControl, _
    ByRef returnedValue)

    Dim numberingMode As String
    Dim separatorText As String

    returnedValue = 0
    On Error GoTo SelectedFinished
    If Documents.Count > 0 Then
        numberingMode = VTEquationNumberingMode(ActiveDocument)
        separatorText = VTEquationNumberingSeparator(ActiveDocument)
    ElseIf Not VTTryReadEquationNumberingPreference( _
       numberingMode, separatorText) Then
        Exit Sub
    End If

    Select Case numberingMode
        Case VT_WORD_NUMBERING_MODE_CHAPTER
            If separatorText = "-" Then
                returnedValue = 2
            Else
                returnedValue = 1
            End If
        Case VT_WORD_NUMBERING_MODE_SECTION
            If separatorText = "-" Then
                returnedValue = 4
            Else
                returnedValue = 3
            End If
        Case Else
            returnedValue = 0
    End Select

SelectedFinished:
End Sub

Public Sub VTWordRibbonApplyNumberingFormat( _
    ByVal control As IRibbonControl, _
    ByVal selectedId As String, _
    ByVal selectedIndex As Integer)

    VTApplyEquationNumberingFormatPreset selectedIndex
End Sub

Public Sub VTWordRibbonNumberingFormatSequence(ByVal control As IRibbonControl)
    VTApplyEquationNumberingFormatPreset 0
End Sub

Public Sub VTWordRibbonNumberingFormatChapterDot(ByVal control As IRibbonControl)
    VTApplyEquationNumberingFormatPreset 1
End Sub

Public Sub VTWordRibbonNumberingFormatChapterDash(ByVal control As IRibbonControl)
    VTApplyEquationNumberingFormatPreset 2
End Sub

Public Sub VTWordRibbonNumberingFormatSectionDot(ByVal control As IRibbonControl)
    VTApplyEquationNumberingFormatPreset 3
End Sub

Public Sub VTWordRibbonNumberingFormatSectionDash(ByVal control As IRibbonControl)
    VTApplyEquationNumberingFormatPreset 4
End Sub

Public Sub VTWordRibbonOpen(ByVal control As IRibbonControl)
    VisualTeX_OpenApplication
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

Private Function VTWordRibbonXmlAttribute( _
    ByVal value As String) As String

    value = Replace$(value, "&", "&amp;")
    value = Replace$(value, """", "&quot;")
    value = Replace$(value, "<", "&lt;")
    value = Replace$(value, ">", "&gt;")
    VTWordRibbonXmlAttribute = value
End Function

Private Function VTWordRibbonReferenceMenuMessage( _
    ByVal labelText As String) As String

    VTWordRibbonReferenceMenuMessage = _
        "<menu xmlns=""" & VTWordRibbonCustomUiNamespace() & """>" & _
        "<button id=""VisualTeX.Mac.Word.CrossReference.Empty""" & _
        " label=""" & VTWordRibbonXmlAttribute(labelText) & """" & _
        " enabled=""false""/>" & _
        "</menu>"
End Function

Private Function VTWordRibbonCustomUiNamespace() As String
    VTWordRibbonCustomUiNamespace = _
        "http:" & Chr$(47) & Chr$(47) & _
        "schemas.microsoft.com/office/2009/07/customui"
End Function

Public Sub VTWordRibbonCrossReferenceMenuContent( _
    ByVal control As IRibbonControl, _
    ByRef returnedValue)

    Dim nativeItems As Variant
    Dim imageItems As Variant
    Dim insertionRange As Range
    Dim itemIndex As Long
    Dim itemText As String
    Dim menuXml As String
    Dim hasNativeItems As Boolean
    Dim hasImageItems As Boolean

    On Error GoTo MenuFailed
    Set VT_WORD_REFERENCE_MENU_DOCUMENT = Nothing
    Set VT_WORD_REFERENCE_MENU_INSERTION_RANGE = Nothing
    If Documents.Count = 0 Then
        returnedValue = VTWordRibbonReferenceMenuMessage( _
            VTUnicodeText(35831, 20808, 25171, 24320, 19968, 20010, _
                87, 111, 114, 100, 32, 25991, 26723))
        Exit Sub
    End If

    Set insertionRange = VTResolveEquationReferenceInsertionRange( _
        Selection.Range.Duplicate)
    If insertionRange Is Nothing Then GoTo MenuFailed
    nativeItems = VTEquationNumberCrossReferenceItems( _
        ActiveDocument, VT_WORD_REFERENCE_KIND_NATIVE)
    imageItems = VTEquationNumberCrossReferenceItems( _
        ActiveDocument, VT_WORD_REFERENCE_KIND_IMAGE)
    hasNativeItems = IsArray(nativeItems)
    hasImageItems = IsArray(imageItems)
    If Not hasNativeItems And Not hasImageItems Then
        returnedValue = VTWordRibbonReferenceMenuMessage( _
            VTUnicodeText(24403, 21069, 25991, 26723, 27809, 26377, _
                21487, 24341, 29992, 30340, 32534, 21495, 20844, 24335))
        Exit Sub
    End If

    Set VT_WORD_REFERENCE_MENU_DOCUMENT = ActiveDocument
    Set VT_WORD_REFERENCE_MENU_INSERTION_RANGE = insertionRange.Duplicate
    menuXml = _
        "<menu xmlns=""" & VTWordRibbonCustomUiNamespace() & """>"
    If hasNativeItems Then
        menuXml = menuXml & _
            "<menu id=""VisualTeX.Mac.Word.CrossReference.Omml""" & _
            " label=""Word OMML " & _
            VTUnicodeText(20844, 24335) & """>"
        For itemIndex = LBound(nativeItems) To UBound(nativeItems)
            itemText = CStr(nativeItems(itemIndex))
            itemText = Replace$(Replace$(Replace$( _
                itemText, vbCr, " "), vbLf, " "), vbTab, " ")
            If Len(itemText) > 180 Then itemText = Left$(itemText, 177) & "..."
            menuXml = menuXml & _
                "<button id=""VisualTeX.Mac.Word.CrossReference.Omml." & _
                CStr(itemIndex) & """ label=""" & _
                VTWordRibbonXmlAttribute(itemText) & """ tag=""" & _
                VT_WORD_REFERENCE_KIND_NATIVE & "|" & CStr(itemIndex) & _
                """ onAction=""VTWordRibbonCrossReferenceItem""/>"
        Next itemIndex
        menuXml = menuXml & "</menu>"
    End If
    If hasImageItems Then
        menuXml = menuXml & _
            "<menu id=""VisualTeX.Mac.Word.CrossReference.Image""" & _
            " label=""" & VTUnicodeText(22270, 29255, 20844, 24335) & _
            """>"
        For itemIndex = LBound(imageItems) To UBound(imageItems)
            itemText = CStr(imageItems(itemIndex))
            itemText = Replace$(Replace$(Replace$( _
                itemText, vbCr, " "), vbLf, " "), vbTab, " ")
            If Len(itemText) > 180 Then itemText = Left$(itemText, 177) & "..."
            menuXml = menuXml & _
                "<button id=""VisualTeX.Mac.Word.CrossReference.Image." & _
                CStr(itemIndex) & """ label=""" & _
                VTWordRibbonXmlAttribute(itemText) & """ tag=""" & _
                VT_WORD_REFERENCE_KIND_IMAGE & "|" & CStr(itemIndex) & _
                """ onAction=""VTWordRibbonCrossReferenceItem""/>"
        Next itemIndex
        menuXml = menuXml & "</menu>"
    End If
    returnedValue = menuXml & "</menu>"
    Exit Sub

MenuFailed:
    Set VT_WORD_REFERENCE_MENU_DOCUMENT = Nothing
    Set VT_WORD_REFERENCE_MENU_INSERTION_RANGE = Nothing
    returnedValue = VTWordRibbonReferenceMenuMessage( _
        VTUnicodeText(26080, 27861, 35835, 21462, 20844, 24335, 21015, 34920))
    Err.Clear
End Sub

Public Sub VTWordRibbonCrossReferenceItem( _
    ByVal control As IRibbonControl)

    Dim itemIndex As Long
    Dim referenceKind As String
    Dim tagParts As Variant
    Dim insertedRange As Range
    Dim insertionRange As Range
    Dim targetDocument As Document

    On Error GoTo InsertFailed
    If VT_WORD_REFERENCE_MENU_DOCUMENT Is Nothing Or _
       VT_WORD_REFERENCE_MENU_INSERTION_RANGE Is Nothing Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "Open the Equation-reference list again before selecting an item."
    End If
    tagParts = Split(CStr(control.Tag), "|")
    If UBound(tagParts) <> 1 Or Not IsNumeric(tagParts(1)) Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The selected Equation-reference item is invalid."
    End If
    referenceKind = LCase$(Trim$(CStr(tagParts(0))))
    If referenceKind <> VT_WORD_REFERENCE_KIND_NATIVE And _
       referenceKind <> VT_WORD_REFERENCE_KIND_IMAGE Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The selected Equation-reference kind is invalid."
    End If
    itemIndex = CLng(tagParts(1))
    Set targetDocument = VT_WORD_REFERENCE_MENU_DOCUMENT
    Set insertionRange = _
        VT_WORD_REFERENCE_MENU_INSERTION_RANGE.Duplicate
    Set VT_WORD_REFERENCE_MENU_DOCUMENT = Nothing
    Set VT_WORD_REFERENCE_MENU_INSERTION_RANGE = Nothing

    targetDocument.Activate
    Set insertedRange = VTInsertEquationNumberReferenceAtRange( _
        insertionRange, itemIndex, True, referenceKind)
    insertedRange.Collapse wdCollapseEnd
    insertedRange.Select
    Exit Sub

InsertFailed:
    Set VT_WORD_REFERENCE_MENU_DOCUMENT = Nothing
    Set VT_WORD_REFERENCE_MENU_INSERTION_RANGE = Nothing
    VTShowError "equation cross-reference", Err.Number, Err.Description
End Sub

Public Sub VTWordRibbonConvertImage(ByVal control As IRibbonControl)
    VisualTeX_ConvertSelectedToImageFormula
End Sub

Public Function VTEnsureVisualTeXImageMacroButton( _
    ByVal formulaShape As InlineShape) As InlineShape

    Dim documentObject As Document
    Dim containingField As Field
    Dim backupDocument As Document
    Dim backupRange As Range
    Dim insertionRange As Range
    Dim restoredShape As InlineShape
    Dim candidateShape As InlineShape
    Dim imageIndex As Long
    Dim sourceStart As Long
    Dim sourceTitle As String
    Dim sourceAlternativeText As String
    Dim sourceWidth As Single
    Dim sourceHeight As Single
    Dim sourceBaseline As Long
    Dim sourceLockAspectRatio As Long
    Dim sourceRemoved As Boolean
    Dim matchCount As Long
    Dim normalizeErrorNumber As Long
    Dim normalizeErrorDescription As String

    ' Word's InlineShape object type is used for both inline formulas and
    ' display formulas placed in their own centred paragraph. The displayMode
    ' metadata and paragraph layout remain authoritative; this routine only
    ' removes the legacy MACROBUTTON field wrapper that can refresh to a visible
    ' fallback glyph and replace the actual formula artwork.
    If formulaShape Is Nothing Then
        Err.Raise vbObjectError + 7593, "VisualTeX", _
            "The VisualTeX image target is missing."
    End If
    Set containingField = VTContainingFieldForInlineShape(formulaShape)
    If containingField Is Nothing Then
        Set VTEnsureVisualTeXImageMacroButton = formulaShape
        Exit Function
    End If
    If Not VTIsVisualTeXImageMacroButtonField(containingField) Or _
       Not VTMacroButtonResultHasOnlyImage(containingField) Then
        Err.Raise vbObjectError + 7593, "VisualTeX", _
            "The VisualTeX image is contained by an incompatible Word field."
    End If

    Set documentObject = formulaShape.Range.Document
    sourceStart = containingField.Code.Start - 1
    sourceTitle = formulaShape.Title
    sourceAlternativeText = formulaShape.AlternativeText
    sourceWidth = formulaShape.Width
    sourceHeight = formulaShape.Height
    sourceBaseline = formulaShape.Range.Font.Position
    sourceLockAspectRatio = formulaShape.LockAspectRatio
    On Error GoTo NormalizeFailed

    Set backupDocument = Documents.Add(Visible:=False)
    Set backupRange = backupDocument.Content.Duplicate
    backupRange.Collapse wdCollapseStart
    backupRange.FormattedText = formulaShape.Range.FormattedText
    If backupDocument.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7593, "VisualTeX", _
            "Word could not back up the legacy VisualTeX image field result."
    End If
    Set backupRange = backupDocument.InlineShapes(1).Range.Duplicate

    containingField.Locked = False
    containingField.Delete
    sourceRemoved = True
    Set containingField = Nothing
    Set formulaShape = Nothing
    Set insertionRange = documentObject.Range( _
        Start:=sourceStart, End:=sourceStart)
    insertionRange.FormattedText = backupRange.FormattedText
    documentObject.Activate

    For Each candidateShape In documentObject.InlineShapes
        If candidateShape.Range.Start = sourceStart Then
            matchCount = matchCount + 1
            Set restoredShape = candidateShape
        End If
    Next candidateShape
    If matchCount <> 1 Or restoredShape Is Nothing Then
        Err.Raise vbObjectError + 7593, "VisualTeX", _
            "Word did not restore exactly one plain VisualTeX image after removing its legacy field."
    End If

    restoredShape.LockAspectRatio = msoFalse
    restoredShape.Width = sourceWidth
    restoredShape.Height = sourceHeight
    restoredShape.LockAspectRatio = sourceLockAspectRatio
    restoredShape.Range.Font.Position = sourceBaseline
    restoredShape.Title = sourceTitle
    restoredShape.AlternativeText = sourceAlternativeText
    Set containingField = VTContainingFieldForInlineShape(restoredShape)
    If Abs(restoredShape.Width - sourceWidth) > 0.1 Or _
       Abs(restoredShape.Height - sourceHeight) > 0.1 Or _
       restoredShape.Range.Font.Position <> sourceBaseline Or _
       StrComp(restoredShape.Title, sourceTitle, vbBinaryCompare) <> 0 Or _
       StrComp(restoredShape.AlternativeText, sourceAlternativeText, _
           vbBinaryCompare) <> 0 Or _
       Not containingField Is Nothing Then
        Err.Raise vbObjectError + 7593, "VisualTeX", _
            "Word changed the VisualTeX image while removing its legacy field wrapper."
    End If

    Set VTEnsureVisualTeXImageMacroButton = restoredShape
    backupDocument.Close SaveChanges:=wdDoNotSaveChanges
    Exit Function

NormalizeFailed:
    normalizeErrorNumber = Err.Number
    normalizeErrorDescription = Err.Description
    On Error Resume Next
    If sourceRemoved And Not documentObject Is Nothing Then
        For imageIndex = documentObject.InlineShapes.Count To 1 Step -1
            If documentObject.InlineShapes(imageIndex).Range.Start = _
               sourceStart Then
                documentObject.InlineShapes(imageIndex).Delete
            End If
        Next imageIndex
        If Not backupRange Is Nothing Then
            Set insertionRange = documentObject.Range( _
                Start:=sourceStart, End:=sourceStart)
            insertionRange.FormattedText = backupRange.FormattedText
        End If
    End If
    If Not backupDocument Is Nothing Then
        backupDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    On Error GoTo 0
    Err.Raise normalizeErrorNumber, "VisualTeX image field migration", _
        normalizeErrorDescription
End Function

Public Sub VTCleanupEmptyVisualTeXImageMacroButtons( _
    ByVal selectedRange As Range)

    Dim paragraphRange As Range
    Dim candidateField As Field
    Dim fieldIndex As Long

    If selectedRange Is Nothing Or VTWordInternalMutationActive() Then Exit Sub
    Set paragraphRange = selectedRange.Paragraphs(1).Range.Duplicate
    For fieldIndex = paragraphRange.Fields.Count To 1 Step -1
        Set candidateField = paragraphRange.Fields(fieldIndex)
        If VTIsVisualTeXImageMacroButtonField(candidateField) Then
            If candidateField.Result.InlineShapes.Count = 0 Then
                candidateField.Delete
            End If
        End If
    Next fieldIndex
End Sub

Public Sub VisualTeX_WriteSelectedDoubleClickScreenBounds()
    Dim formulaShape As InlineShape
    Dim targetRange As Range
    Dim screenLeft As Long
    Dim screenTop As Long
    Dim screenWidth As Long
    Dim screenHeight As Long
    Dim statusPath As String
    Dim boundsErrorNumber As Long
    Dim boundsErrorDescription As String

    statusPath = VTApplicationSupportRoot() & _
        "/double-click-screen-bounds.txt"
    On Error GoTo Failed
    If Documents.Count = 0 Or Selection Is Nothing Then
        Err.Raise vbObjectError + 7593, "VisualTeX", _
            "There is no selected Word formula for double-click hit testing."
    End If
    If Not VTTryVisualTeXMetadataShapeAtSelection( _
       Selection, formulaShape) Then
        Err.Raise vbObjectError + 7593, "VisualTeX", _
            "The current Word selection does not resolve to a VisualTeX image formula."
    End If
    Set targetRange = formulaShape.Range.Duplicate
    ActiveWindow.GetPoint _
        screenLeft, screenTop, screenWidth, screenHeight, targetRange
    If screenWidth <= 0 Or screenHeight <= 0 Then
        Err.Raise vbObjectError + 7593, "VisualTeX", _
            "Word returned invalid formula screen bounds."
    End If
    VTWriteTextAtomic statusPath, _
        "PASS|" & CStr(screenLeft) & "|" & CStr(screenTop) & "|" & _
        CStr(screenWidth) & "|" & CStr(screenHeight) & vbLf
    Exit Sub

Failed:
    boundsErrorNumber = Err.Number
    boundsErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic statusPath, _
        "FAIL|" & CStr(boundsErrorNumber) & "|" & _
        Replace$(Replace$(boundsErrorDescription, vbCr, " "), _
            vbLf, " ") & vbLf
    On Error GoTo 0
End Sub

Public Sub VisualTeX_WriteSelectedDoubleClickTargetScreenBounds()
    Dim formulaShape As InlineShape
    Dim nativeBookmark As Bookmark
    Dim nativeMath As OMath
    Dim targetRange As Range
    Dim screenLeft As Long
    Dim screenTop As Long
    Dim screenWidth As Long
    Dim screenHeight As Long
    Dim statusPath As String
    Dim boundsErrorNumber As Long
    Dim boundsErrorDescription As String

    statusPath = VTApplicationSupportRoot() & _
        "/double-click-target-screen-bounds.txt"
    On Error GoTo Failed
    If Documents.Count = 0 Or Selection Is Nothing Then
        Err.Raise vbObjectError + 7593, "VisualTeX", _
            "There is no selected Word formula for double-click target testing."
    End If
    Set formulaShape = VTVisualTeXInlineShapeAtSelection(Selection)
    If Not formulaShape Is Nothing Then
        Set targetRange = formulaShape.Range.Duplicate
    ElseIf VTTryFindNativeFormulaBookmarkLocally( _
       Selection.Range, nativeBookmark) Then
        Set targetRange = nativeBookmark.Range.Duplicate
    ElseIf VTTryResolveSingleNativeMath(Selection.Range, nativeMath) Then
        Set targetRange = nativeMath.Range.Duplicate
    End If
    If targetRange Is Nothing Then
        Err.Raise vbObjectError + 7593, "VisualTeX", _
            "The current Word selection is not a formula target."
    End If
    ActiveWindow.GetPoint _
        screenLeft, screenTop, screenWidth, screenHeight, targetRange
    If screenWidth <= 0 Or screenHeight <= 0 Then
        Err.Raise vbObjectError + 7593, "VisualTeX", _
            "Word returned invalid formula target screen bounds."
    End If
    VTWriteTextAtomic statusPath, _
        "PASS|" & CStr(screenLeft) & "|" & CStr(screenTop) & "|" & _
        CStr(screenWidth) & "|" & CStr(screenHeight) & vbLf
    Exit Sub

Failed:
    boundsErrorNumber = Err.Number
    boundsErrorDescription = Err.Description
    On Error Resume Next
    VTWriteTextAtomic statusPath, _
        "FAIL|" & CStr(boundsErrorNumber) & "|" & _
        Replace$(Replace$(boundsErrorDescription, vbCr, " "), _
            vbLf, " ") & vbLf
    On Error GoTo 0
End Sub

Public Sub VTEnsureImageMacroMigrationScheduled()
    If VTWordVbeBuildTemplateActive() Then Exit Sub
    On Error GoTo ScheduleFailed
    If VT_WORD_IMAGE_MACRO_MIGRATION_SCHEDULED Or _
       VT_WORD_IMAGE_MACRO_MIGRATION_RUNNING Or _
       Documents.Count = 0 Then Exit Sub
    VT_WORD_IMAGE_MACRO_MIGRATION_SCHEDULED = True
    VTScheduleWordIdleTasks 1
    Exit Sub

ScheduleFailed:
    VT_WORD_IMAGE_MACRO_MIGRATION_SCHEDULED = False
End Sub

Public Sub VisualTeX_MigrateImageMacroButtons()
    Dim documentObject As Document

    VT_WORD_IMAGE_MACRO_MIGRATION_SCHEDULED = False
    If VT_WORD_IMAGE_MACRO_MIGRATION_RUNNING Then Exit Sub
    VT_WORD_IMAGE_MACRO_MIGRATION_RUNNING = True
    On Error Resume Next
    For Each documentObject In Documents
        VTMigrateDocumentImageMacroButtons documentObject
        Err.Clear
    Next documentObject
    On Error GoTo 0
    VT_WORD_IMAGE_MACRO_MIGRATION_RUNNING = False
End Sub

Public Sub VTMigrateOpenedDocumentImageMacroButtons( _
    ByVal documentObject As Word.Document)

    If documentObject Is Nothing Or VTWordVbeBuildTemplateActive() Then Exit Sub
    VTMigrateDocumentImageMacroButtons documentObject
End Sub

' Opening a document WPS has saved (it records its build in the custom
' property KSOProductBuildVer, which Word keeps): rebuild the equation numbers
' WPS damaged before they are read. Other documents are left alone without
' scanning their fields.
Public Sub VTRepairOpenedWpsDocument(ByVal documentObject As Word.Document)
    Dim savedByWps As Boolean
    Dim mutationStarted As Boolean
    Dim previousScreenUpdating As Boolean

    If documentObject Is Nothing Or VTWordVbeBuildTemplateActive() Then Exit Sub
    If VT_WORD_NUMBER_UPDATE_RUNNING Then Exit Sub
    On Error Resume Next
    savedByWps = Len(CStr(documentObject.CustomDocumentProperties( _
        "KSOProductBuildVer").Value)) > 0
    On Error GoTo Finished
    If Not savedByWps Then Exit Sub
    If documentObject.ProtectionType <> wdNoProtection Then Exit Sub
    If documentObject.ReadOnly Then Exit Sub
    If Not VTHasWpsDamagedEquationNumber(documentObject) Then Exit Sub
    previousScreenUpdating = Application.ScreenUpdating
    VT_WORD_NUMBER_UPDATE_RUNNING = True
    VTBeginWordInternalMutation
    mutationStarted = True
    Application.ScreenUpdating = False
    VTV2RestyleNumbers documentObject

Finished:
    If mutationStarted Then
        Application.ScreenUpdating = previousScreenUpdating
        VTEndWordInternalMutation
        VT_WORD_NUMBER_UPDATE_RUNNING = False
    End If
End Sub

Public Sub VisualTeX_EditSelected()
    Dim selectedShape As InlineShape
    Dim nativeBookmark As Bookmark

    On Error GoTo Failed

    VTRequireWritableWordDocument
    Set selectedShape = VTVisualTeXInlineShapeAtSelection(Selection)
    If Not selectedShape Is Nothing Then
        VTWordEditInlineShape selectedShape
    Else
        If Not VTWordOpenCopiedNativeSession(Selection.Range) Then
            Err.Raise vbObjectError + 7461, "VisualTeX", _
                "Select one VisualTeX formula image or native equation."
        End If
    End If
    Exit Sub

Failed:
    VTShowError "Word edit", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_DoubleClickEditSelected()
    On Error Resume Next
    VTHandleWordBeforeDoubleClick Selection
    On Error GoTo 0
End Sub

Public Sub VisualTeX_EditImageField()
    Dim editErrorNumber As Long
    Dim editErrorDescription As String

    On Error GoTo Failed
    VTBeginWordDoubleClickTraceContext
    VTTraceWordDoubleClick "macrobutton-enter", Selection, ""
    If Not VTDispatchVisualTeXImageEditAtSelection(Selection) Then
        VTTraceWordDoubleClick "macrobutton-skip", Selection, _
            "reason=no-visualtex-image"
    End If
    VTEndWordDoubleClickTraceContext
    Exit Sub

Failed:
    editErrorNumber = Err.Number
    editErrorDescription = Err.Description
    VTTraceWordDoubleClick _
        "macrobutton-error", Selection, _
        "error=" & CStr(editErrorNumber) & ":" & editErrorDescription
    VTEndWordDoubleClickTraceContext
    VTShowError _
        "Word image formula edit", editErrorNumber, editErrorDescription
End Sub

Public Sub VisualTeX_EditSelectedImageFromNativeMonitor()
    Dim selectedShape As InlineShape

    On Error GoTo MonitorFailed
    VTBeginWordDoubleClickTraceContext
    VTTraceWordDoubleClick "native-monitor-enter", Selection, ""
    If Not VTTryVisualTeXMetadataShapeAtSelection( _
       Selection, selectedShape) Then
        VTTraceWordDoubleClick _
            "native-monitor-skip", Selection, "reason=no-metadata-image"
        VTEndWordDoubleClickTraceContext
        Exit Sub
    End If

    ' The macOS native monitor owns task-pane cleanup because a global mouse
    ' monitor cannot cancel Word's native second-click action. This VBA entry
    ' performs only the strict metadata-resolved image edit.
    VTWordEditInlineShape selectedShape
    VTEndWordDoubleClickTraceContext
    Exit Sub

MonitorFailed:
    VTTraceWordDoubleClick _
        "native-monitor-error", Selection, _
        "error=" & CStr(Err.Number) & ":" & Err.Description
    VTEndWordDoubleClickTraceContext
    ' The native monitor observes ordinary pictures too. It is deliberately a
    ' silent, image-only route and never falls through to OMML or Word commands.
End Sub

Public Sub FormatPicture()
    If VTInterceptVisualTeXPictureFormatCommand("format-picture") Then Exit Sub

    ' Ordinary pictures do not carry VisualTeX's compressed AlternativeText
    ' metadata. Preserve Word's native Format Picture behavior exactly.
    On Error Resume Next
    WordBasic.FormatPicture
    On Error GoTo 0
End Sub

Public Sub FormatDrawingObject()
    If VTInterceptVisualTeXPictureFormatCommand( _
       "format-drawing-object") Then Exit Sub

    On Error Resume Next
    WordBasic.FormatDrawingObject
    On Error GoTo 0
End Sub

Public Sub WW7_FormatDrawingObject()
    If VTInterceptVisualTeXPictureFormatCommand( _
       "ww7-format-drawing-object") Then Exit Sub

    On Error Resume Next
    Err.Clear
    WordBasic.WW7_FormatDrawingObject
    If Err.Number <> 0 Then
        Err.Clear
        WordBasic.FormatDrawingObject
    End If
    On Error GoTo 0
End Sub

Public Sub FormatObjectCore()
    If VTInterceptVisualTeXPictureFormatCommand( _
       "format-object-core") Then Exit Sub

    On Error Resume Next
    Err.Clear
    WordBasic.FormatObjectCore
    If Err.Number <> 0 Then
        Err.Clear
        WordBasic.FormatDrawingObject
    End If
    On Error GoTo 0
End Sub

Public Function VTHandleWordBeforeDoubleClick( _
    ByVal selected As Word.Selection) As Boolean

    Dim selectedShape As InlineShape
    Dim nativeBookmark As Bookmark
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean

    ' Both the real Application event and the regression macro use this one
    ' target resolver. It must remain a strict no-op for ordinary Word content.
    VTBeginWordDoubleClickTraceContext
    On Error GoTo IgnoreDoubleClick
    VTTraceWordDoubleClick "handler-enter", selected, ""
    If Documents.Count = 0 Then
        VTTraceWordDoubleClick "handler-skip", selected, "reason=no-document"
        GoTo HandlerFinished
    End If
    If selected Is Nothing Then
        VTTraceWordDoubleClick "handler-skip", selected, "reason=no-selection"
        GoTo HandlerFinished
    End If
    If VTWordInternalMutationActive() Then
        VTTraceWordDoubleClick "handler-skip", selected, "reason=internal-mutation"
        GoTo HandlerFinished
    End If
    If VTTryVisualTeXMetadataShapeAtDoubleClick( _
       selected, selectedShape) Then
        VTHandleWordBeforeDoubleClick = True
        VTTraceWordDoubleClick _
            "handler-image-metadata-resolved", selected, _
            VTInlineShapeTraceSummary(selectedShape)
        If Not VTTryResolveVisualTeXInlineShapeReference( _
           selectedShape, formulaId, displayMode, numbered, _
           encodedMetadata, metadataNeedsWrite, formatNeedsWrite) Then
            VTTraceWordDoubleClick _
                "handler-image-metadata-invalid", selected, _
                VTInlineShapeTraceSummary(selectedShape)
            VTShowError _
                "Word image formula edit", vbObjectError + 7400, _
                "The selected VisualTeX image metadata cannot be restored."
            GoTo HandlerFinished
        End If
        VTWordOpenResolvedInlineShape _
            selectedShape, formulaId, displayMode, numbered, _
            encodedMetadata, metadataNeedsWrite, formatNeedsWrite, False
        VTTraceWordDoubleClick "handler-image-edit-dispatched", selected, ""
        GoTo HandlerFinished
    End If
    VTTraceWordDoubleClick "handler-image-not-found", selected, ""
    VTHandleWordBeforeDoubleClick = VTWordOpenCopiedNativeSession(selected.Range)
    VTTraceWordDoubleClick "handler-native-edit-dispatched", selected, ""

HandlerFinished:
    VTEndWordDoubleClickTraceContext
    Exit Function

IgnoreDoubleClick:
    VTTraceWordDoubleClick _
        "handler-error", selected, _
        "error=" & CStr(Err.Number) & ":" & Err.Description
    ' Preserve Word's native double-click behavior for every non-VisualTeX
    ' target, including empty space and ordinary document text.
    VTHandleWordBeforeDoubleClick = False
    Resume HandlerFinished
End Function

Public Sub VisualTeX_EditInlineShape(ByVal selectedShape As InlineShape)
    On Error GoTo Failed
    VTRequireWritableWordDocument
    VTWordEditInlineShape selectedShape
    Exit Sub
Failed:
    VTShowError "Word edit", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_EditNativeSelection(ByVal selectedRange As Range)
    Dim nativeBookmark As Bookmark

    On Error GoTo Failed
    VTRequireWritableWordDocument
    Set nativeBookmark = VTFindNativeFormulaBookmark(selectedRange, False)
    If Not nativeBookmark Is Nothing Then
        VTWordEditNativeBookmark nativeBookmark
    ElseIf Not VTWordOpenCopiedNativeSession(selectedRange) Then
        Err.Raise vbObjectError + 7461, "VisualTeX", _
            "Select one VisualTeX native Word equation."
    End If
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
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean

    On Error GoTo InvalidShape
    VTIsVisualTeXInlineShape = _
        VTTryResolveVisualTeXInlineShapeReference( _
            selectedShape, formulaId, displayMode, numbered, _
            encodedMetadata, metadataNeedsWrite, formatNeedsWrite)
    Exit Function
InvalidShape:
    VTIsVisualTeXInlineShape = False
End Function

Public Function VTVisualTeXInlineShapeAtSelection( _
    ByVal selected As Selection, _
    Optional ByVal repairShapeProperties As Boolean = True) As InlineShape

    Dim formulaShape As InlineShape
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean

    If VTTryResolveVisualTeXInlineShapeAtSelection( _
       selected, formulaShape, formulaId, displayMode, numbered, _
       encodedMetadata, metadataNeedsWrite, formatNeedsWrite, _
       repairShapeProperties) Then
        Set VTVisualTeXInlineShapeAtSelection = formulaShape
    End If
End Function

Public Sub VisualTeX_ConvertSelectedToNativeEquation()
    Dim selectedShape As InlineShape
    Dim targetDocument As Document
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim formulaReference As String
    Dim sourceStart As Long
    Dim hasStableIdentity As Boolean
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean
    Dim conversionStage As String

    On Error GoTo Failed

    conversionStage = "validate-document"
    VTRequireWritableWordDocument
    conversionStage = "resolve-image-selection"
    Set selectedShape = VTVisualTeXInlineShapeAtSelection(Selection)
    If selectedShape Is Nothing Then
        Err.Raise vbObjectError + 7428, "VisualTeX", _
            "Select exactly one VisualTeX formula image to convert to Word OMML."
    End If
    Set targetDocument = selectedShape.Range.Document
    sourceStart = selectedShape.Range.Start
    If VTTryResolveVisualTeXInlineShapeReference( _
       selectedShape, formulaId, displayMode, numbered, encodedMetadata, _
       metadataNeedsWrite, formatNeedsWrite) Then
        hasStableIdentity = True
        formulaReference = VTFormulaReference( _
            formulaId, displayMode, numbered)
        If VTShouldForkWordImageFormulaCopy( _
           selectedShape.Range.Document, selectedShape, formulaId, _
           encodedMetadata) Then
            ' A copied image must receive its own metadata and formula id before
            ' conversion. Route it through the renderer fork transaction rather
            ' than the formula-id keyed fast cache, which belongs to the source.
            conversionStage = "fork-copied-image-to-native"
            VTWordEditInlineShape selectedShape, True, "imageToNative"
            Exit Sub
        End If
    End If
    ' A copied image must obtain a fresh formula identity before conversion.
    ' Its Title/AlternativeText still point at the source formula, so neither
    ' the warm-cache fast path nor the legacy direct converter may run with the
    ' inherited id. Use the renderer-backed edit transaction: Rust forks the
    ' metadata/line ids and Word commits the native result to this exact copy.
    If VTIsCopiedWordImageFormula(selectedShape) Then
        conversionStage = "fork-copied-image-to-native"
        VTWordEditInlineShape selectedShape, True, "imageToNative"
        Exit Sub
    End If

    ' A current VisualTeX image already carries the structural OMML payload and
    ' normally has a durable native DOCX cache. Convert it entirely inside Word
    ' to avoid waking the hidden renderer and paying two AppleEvent round trips.
    ' Older images without that cache retain the validated hidden-renderer
    ' fallback, so compatibility and all existing error UI remain unchanged.
    conversionStage = "convert-image-to-native-fast"
    If VTWordConvertInlineShapeToNativeFast(selectedShape) Then Exit Sub
    ' A failed fast attempt can insert/delete fields and then Undo. Word returns
    ' a new InlineShape wrapper for the restored picture, so the object captured
    ' above is no longer safe even though the visible image is back. Resolve the
    ' formula again from its immutable id and metadata before compatibility work.
    If hasStableIdentity Then
        conversionStage = "re-resolve-image-after-fast-path"
        Set selectedShape = VTResolveImageFormulaByIdentity( _
            targetDocument, formulaId, encodedMetadata, _
            formulaReference, sourceStart)
    Else
        conversionStage = "re-resolve-image-selection-after-fast-path"
        Set selectedShape = VTVisualTeXInlineShapeAtSelection(Selection)
        If selectedShape Is Nothing Then
            Err.Raise vbObjectError + 7428, "VisualTeX", _
                "Word could not restore the selected VisualTeX image after the fast conversion attempt."
        End If
    End If
    conversionStage = "convert-image-to-native-compatible"
    VTWordConvertInlineShapeToNativeEquation selectedShape
    Exit Sub

Failed:
    VTShowError "Word native equation conversion", Err.Number, _
        conversionStage & ": " & Err.Description
End Sub

Public Sub VisualTeX_ConvertSelectedToImageFormula()
    Dim nativeBookmark As Bookmark

    On Error GoTo Failed
    VTRequireWritableWordDocument
    Set nativeBookmark = VTFindNativeFormulaBookmark(Selection.Range, False)
    If nativeBookmark Is Nothing Then
        If VTWordOpenCopiedNativeSession( _
           Selection.Range, False, "nativeToImage") Then Exit Sub
        VTStartWordFormulaRestore "selection", "omml", "image"
    Else
        ' Managed OMML created from a current VisualTeX image already has a
        ' verified native signature, SVG cache and fractional scale state. Use
        ' that complete in-Word replacement first: the hidden nativeToImage URL
        ' can be accepted by the helper process without reaching the resident
        ' helper, leaving the Session at "created" and making the Ribbon appear
        ' to do nothing. Older or incomplete formulas still fall back to the
        ' renderer so their geometry can be regenerated.
        If VTWordConvertNativeBookmarkToImageFast(nativeBookmark) Then Exit Sub
        VTWordOpenNativeSession nativeBookmark, False, "nativeToImage"
    End If
    Exit Sub

Failed:
    VTShowError "Word image formula conversion", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_UpdateEquationNumbers()
    Dim updated As Long

    If VT_WORD_NUMBER_UPDATE_RUNNING Then Exit Sub
    On Error GoTo Failed
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7401, "VisualTeX", "Open a Word document first."
    End If
    updated = VTUpdateEquationNumbersCore(ActiveDocument, True)
    ' A successful refresh must return control immediately. A modal MsgBox can
    ' remain behind another Office/VisualTeX window and make Word look frozen.
    Application.StatusBar = "VisualTeX: updated " & CStr(updated) & _
        " equation numbers."
    Exit Sub

Failed:
    VTShowError "equation numbering", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_OpenEquationCrossReference()
    Dim crossReferenceItems As Variant
    Dim selectedIndex As Variant
    Dim promptText As String
    Dim itemText As String
    Dim itemIndex As Long
    Dim itemCount As Long
    Dim displayCount As Long
    Dim insertedRange As Range
    Dim insertionRange As Range

    On Error GoTo Failed
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7401, "VisualTeX", "Open a Word document first."
    End If

    ' Freeze and normalize the user's target before building the picker. Opening
    ' the list remains read-only; the only recovery allowed here is the local
    ' helper-to-continuation correction when Word itself left Selection inside a
    ' VisualTeX numbering helper after the preceding formula mutation.
    Set insertionRange = VTResolveEquationReferenceInsertionRange( _
        Selection.Range.Duplicate)
    If insertionRange Is Nothing Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The Equation cross-reference insertion point could not be resolved."
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
    If Len(Trim$(CStr(selectedIndex))) = 0 Then
        insertionRange.Select
        Exit Sub
    End If
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
        insertionRange, itemIndex, True)
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
    Dim callbackErrorNumber As Long
    Dim callbackErrorDescription As String

    VTWordPerformanceStart
    On Error GoTo Failed
    VTBeginNumberingResolveScope
    sessionId = VTReadActiveSessionId(VT_WORD_HOST)
    VTTraceWordSession sessionId, "callback-enter", ""
    Set dispatch = VTReadDispatch(sessionId)
    VTWordPerformanceEnable dispatch
    actionName = CStr(dispatch("action"))
    hostName = CStr(dispatch("host"))
    If hostName <> VT_WORD_HOST Then
        Err.Raise vbObjectError + 7402, "VisualTeX", "The active VisualTeX dispatch is not for Word."
    End If

    Select Case actionName
        Case "commit": VTCommitWordDispatch sessionId, dispatch
        Case "cancel": VTCancelWordDispatch sessionId, dispatch
        Case "latexRedrawPreflight": _
            VTResolveWordLatexRedrawFontsDispatch sessionId, dispatch
        Case "documentCommit": _
            VTCommitWordDocumentImportDispatch sessionId, dispatch
        Case "documentCancel": _
            VTCancelWordDocumentImportDispatch dispatch
        Case Else
            Err.Raise vbObjectError + 7403, "VisualTeX", "The VisualTeX Word dispatch action is invalid."
    End Select
    VTWordPerformanceMark "callback-complete"
    VTEndNumberingResolveScope
    VTWordPerformanceFlush sessionId
    Exit Sub

Failed:
    callbackErrorNumber = Err.Number
    callbackErrorDescription = Err.Description
    On Error Resume Next
    VTEndNumberingResolveScope
    If Len(sessionId) > 0 Then
        ' documentCommit already records the exact inner import/redraw stage.
        ' Preserve that trace instead of replacing it with this callback wrapper.
        If actionName <> "documentCommit" Then
            VTWriteWordFailureTrace _
                sessionId, "document-import-callback", _
                callbackErrorNumber, callbackErrorDescription
        End If
        VTWordPerformanceMark "callback-failed"
        VTWordPerformanceFlush _
            sessionId, callbackErrorNumber, callbackErrorDescription
    End If
    On Error GoTo 0
    If VTReportCallbackFailure(sessionId, callbackErrorNumber, callbackErrorDescription) Then Exit Sub
    Err.Raise callbackErrorNumber, "VisualTeX Word callback", _
        callbackErrorDescription
End Sub

Public Sub VTResetWordImageObservationCache()
    ' A reopened file may have changed outside this Word session while keeping
    ' the same formula ids and picture frame dimensions.
    Set VT_WORD_IMAGE_CROP_CACHE = Nothing
    VT_WORD_SELECTED_IMAGE_FINGERPRINT = ""
    VT_WORD_TYPOGRAPHY_FINGERPRINT = ""
    VT_WORD_DOCUMENT_TYPOGRAPHY_FINGERPRINT = ""
End Sub

Public Sub VTSynchronizeWordImageTypography(ByVal contextRange As Range)
    Dim paragraph As Paragraph
    Dim shape As InlineShape
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim fingerprint As String
    Dim mutationStarted As Boolean
    Dim numberRange As Range
    Dim keptNumberFormat As Boolean
    On Error GoTo Finished
    If contextRange Is Nothing Or VTWordInternalMutationActive() Then Exit Sub
    If contextRange.Document.ReadOnly Or _
       contextRange.Document.ProtectionType <> wdNoProtection Then Exit Sub
    fingerprint = VTImageTypographyFingerprint(contextRange)
    If fingerprint = VT_WORD_TYPOGRAPHY_FINGERPRINT Then Exit Sub
    ' Re-align pictures to their numbers as the user formatted them.
    keptNumberFormat = VT_WORD_KEEP_NUMBER_FORMAT
    VT_WORD_KEEP_NUMBER_FORMAT = True
    For Each paragraph In contextRange.Paragraphs
        For Each shape In paragraph.Range.InlineShapes
            If VTTryWordImageTypographyReference(shape, formulaId, displayMode, numbered) Then
                If Not mutationStarted Then
                    VTBeginWordInternalMutation
                    mutationStarted = True
                End If
                ' Actual synchronization may restore legacy carriers; queries
                ' and fingerprints use the read-only default.
                VTTryWordImageTypographyReference shape, formulaId, displayMode, numbered, True
                If displayMode = "inline" Then
                    ' Inline formulas keep their Position until they are edited
                    ' or resized, or the user runs Update Equation Format.
                ElseIf numbered Then
                    Set numberRange = VTStaticImageEquationNumberRange(shape.Range, formulaId)
                    If Not numberRange Is Nothing Then
                        VTApplyNumberedImageFormulaVerticalAlignment shape, numberRange, _
                            VTVisibleEquationNumberFontSize(contextRange.Document), "Word typography"
                    Else
                        VTApplyWordImagePosition shape, VTExpectedImageFormulaPosition( _
                            shape, VTVisibleEquationNumberFontSize(contextRange.Document), "Word typography")
                    End If
                Else
                    VTApplyWordImagePosition shape, VTExpectedImageFormulaPosition( _
                        shape, VTVisibleEquationNumberFontSize(contextRange.Document), "Word typography")
                End If
            End If
        Next shape
    Next paragraph
    VT_WORD_TYPOGRAPHY_FINGERPRINT = VTImageTypographyFingerprint(contextRange)
Finished:
    VT_WORD_KEEP_NUMBER_FORMAT = keptNumberFormat
    If mutationStarted Then VTEndWordInternalMutation
End Sub

Public Sub VisualTeX_SynchronizeSelectedImageFormulaSize( _
    ByVal selected As Selection)

    Dim formulaShape As InlineShape
    Dim internalMutationStarted As Boolean
    Dim fingerprint As String
    Dim legacyField As Field

    On Error GoTo SynchronizeFinished
    If selected Is Nothing Or VTWordInternalMutationActive() Then Exit Sub
    If selected.Range.Document.ReadOnly Or _
       selected.Range.Document.ProtectionType <> wdNoProtection Then Exit Sub
    ' The selection's font size is part of it: changing it in Word's font size
    ' box resizes the formula, and changes nothing else this compares.
    fingerprint = CStr(selected.Range.Start) & ":" & CStr(selected.Range.End) & _
        ":" & VTSelectionFontSizeText(selected) & ":" & VTImageTypographyFingerprint(selected.Range)
    If fingerprint = VT_WORD_SELECTED_IMAGE_FINGERPRINT Then GoTo SynchronizeFinished
    Set formulaShape = VTVisualTeXInlineShapeAtSelection(selected)
    If formulaShape Is Nothing Then GoTo RepairParagraphTypography
    Set legacyField = VTContainingFieldForInlineShape(formulaShape)
    If Not legacyField Is Nothing Then
        VTBeginWordInternalMutation
        internalMutationStarted = True
        Set formulaShape = _
            VTEnsureVisualTeXImageMacroButton(formulaShape)
        VTEndWordInternalMutation
        internalMutationStarted = False
    End If
    VTSynchronizeWordImageFormulaShape formulaShape

RepairParagraphTypography:
    VTSynchronizeWordImageTypography selected.Range
    VT_WORD_SELECTED_IMAGE_FINGERPRINT = _
        CStr(selected.Range.Start) & ":" & CStr(selected.Range.End) & ":" & _
        VTSelectionFontSizeText(selected) & ":" & VTImageTypographyFingerprint(selected.Range)
SynchronizeFinished:
    If internalMutationStarted Then VTEndWordInternalMutation
    VTEnsureImageSizeWatchScheduled
End Sub

Public Sub VisualTeX_RunWordIdleTasks()
    ' Word has one pending OnTime slot. Every requester therefore uses this
    ' callback; the earliest schedule drains all outstanding work. Individual
    ' tasks retain their guards and reschedule through this same dispatcher.
    If VT_WORD_IDLE_TASKS_RUNNING Then Exit Sub
    VT_WORD_IDLE_TASKS_RUNNING = True
    VT_WORD_IDLE_TASK_DUE_AT = 0
    On Error Resume Next
    If VT_WORD_APPLICATION_PREWARM_SCHEDULED Then _
        VisualTeX_PrewarmWordApplication
    If VT_WORD_IMAGE_MACRO_MIGRATION_SCHEDULED Then _
        VisualTeX_MigrateImageMacroButtons
    If VT_WORD_ORPHAN_WATCH_SCHEDULED Then _
        VisualTeX_WatchOrphanedNumberedDisplay
    If VT_WORD_IMAGE_SIZE_WATCH_SCHEDULED Then _
        VisualTeX_WatchSelectedImageFormulaSize
    On Error GoTo 0
    VT_WORD_IDLE_TASKS_RUNNING = False
End Sub

Public Sub VisualTeX_WatchSelectedImageFormulaSize()
    Dim fingerprint As String
    VT_WORD_IMAGE_SIZE_WATCH_SCHEDULED = False
    If VT_WORD_IMAGE_SIZE_WATCH_RUNNING Then Exit Sub

    VT_WORD_IMAGE_SIZE_WATCH_RUNNING = True
    On Error Resume Next
    If Documents.Count > 0 And Not VTWordInternalMutationActive() Then
        VisualTeX_SynchronizeSelectedImageFormulaSize Selection
        ' While a formula is selected the watch runs every second; the whole
        ' document is still compared only every VT_WORD_IDLE_WATCH_SECONDS.
        VT_WORD_IMAGE_WATCH_SECONDS_SINCE_SCAN = _
            VT_WORD_IMAGE_WATCH_SECONDS_SINCE_SCAN + VT_WORD_IMAGE_WATCH_LAST_DELAY
        If VT_WORD_IMAGE_WATCH_SECONDS_SINCE_SCAN >= VT_WORD_IDLE_WATCH_SECONDS Then
        VT_WORD_IMAGE_WATCH_SECONDS_SINCE_SCAN = 0
#If VT_WORD_DIAGNOSTICS Then
        If Not VT_WORD_TYPOGRAPHY_REGRESSION_RUNNING Then
#End If
            fingerprint = VTDocumentImageTypographyFingerprint(ActiveDocument)
            If fingerprint <> VT_WORD_DOCUMENT_TYPOGRAPHY_FINGERPRINT Then
                VT_WORD_TYPOGRAPHY_FINGERPRINT = ""
                VTSynchronizeWordImageTypography ActiveDocument.Content
                VT_WORD_DOCUMENT_TYPOGRAPHY_FINGERPRINT = _
                    VTDocumentImageTypographyFingerprint(ActiveDocument)
            End If
#If VT_WORD_DIAGNOSTICS Then
        End If
#End If
        End If
    End If
    On Error GoTo 0
    VT_WORD_IMAGE_SIZE_WATCH_RUNNING = False
    VTEnsureImageSizeWatchScheduled
End Sub

Public Sub VTInvalidateWordImageFontSizeControl()
    Dim formulaCount As Long
    Dim mixedSizes As Boolean
    Dim fontSizePt As Double
    Dim controlState As String
    On Error GoTo Finished
    If VT_WORD_RIBBON Is Nothing Then Exit Sub
    VTWordSelectedImageFormulaState formulaCount, mixedSizes, fontSizePt
    controlState = CStr(formulaCount) & ":" & CStr(mixedSizes) & ":" & _
        VTJsonNumber(fontSizePt)
    If controlState = VT_WORD_IMAGE_FONT_CONTROL_STATE Then Exit Sub
    VT_WORD_RIBBON.InvalidateControl VT_WORD_IMAGE_FONT_CONTROL_ID
    VT_WORD_IMAGE_FONT_CONTROL_STATE = controlState
Finished:
    On Error GoTo 0
End Sub

Public Sub VTInvalidateWordEquationNumberingFormatControl()
    On Error Resume Next
    If Not VT_WORD_RIBBON Is Nothing Then
        VT_WORD_RIBBON.InvalidateControl VT_WORD_NUMBERING_FORMAT_CONTROL_ID
    End If
    On Error GoTo 0
End Sub

Public Sub VTWordRibbonGetImageFontSizeItemCount( _
    ByVal control As IRibbonControl, _
    ByRef returnedValue)

    returnedValue = VTFormulaFontPresetCount() + 1
End Sub

Public Sub VTWordRibbonGetImageFontSizeItemLabel( _
    ByVal control As IRibbonControl, _
    ByVal itemIndex As Integer, _
    ByRef returnedValue)

    Dim formulaCount As Long
    Dim mixedSizes As Boolean
    Dim fontSizePt As Double

    On Error GoTo LabelFailed
    If itemIndex = 0 Then
        If Not VTWordSelectedImageFormulaState( _
           formulaCount, mixedSizes, fontSizePt) Then
            returnedValue = _
                VTUnicodeText(35831, 36873, 25321, 22270, 29255, 20844, 24335)
        ElseIf mixedSizes Then
            returnedValue = _
                VTUnicodeText(24403, 21069) & ": " & _
                VTUnicodeText(28151, 21512, 23383, 21495)
        Else
            returnedValue = _
                VTUnicodeText(24403, 21069) & ": " & _
                VTFormulaFontSizeDisplayLabel(fontSizePt)
        End If
    Else
        returnedValue = VTFormulaFontPresetLabel(itemIndex - 1)
    End If
    Exit Sub

LabelFailed:
    returnedValue = VTUnicodeText(23383, 21495)
    Err.Clear
End Sub

Public Sub VTWordRibbonGetImageFontSizeSelectedIndex( _
    ByVal control As IRibbonControl, _
    ByRef returnedValue)

    Dim formulaCount As Long
    Dim mixedSizes As Boolean
    Dim fontSizePt As Double
    Dim presetIndex As Long

    returnedValue = 0
    On Error GoTo SelectedFinished
    If Not VTWordSelectedImageFormulaState( _
       formulaCount, mixedSizes, fontSizePt) Then Exit Sub
    If mixedSizes Then Exit Sub
    presetIndex = VTFormulaFontPresetIndex(fontSizePt)
    If presetIndex >= 0 Then returnedValue = presetIndex + 1

SelectedFinished:
End Sub

Public Sub VTWordRibbonApplyImageFontSizePreset( _
    ByVal control As IRibbonControl, _
    ByVal selectedId As String, _
    ByVal selectedIndex As Integer)

    On Error GoTo Failed
    If selectedIndex <= 0 Then Exit Sub
    VTApplyWordImageFontSizePresetToSelection _
        VTFormulaFontPresetSize(selectedIndex - 1)
    Exit Sub

Failed:
    VTShowError "Word image formula font size", Err.Number, Err.Description
End Sub

Public Sub VTWordRibbonGetConversionFontSizeItemCount( _
    ByVal control As IRibbonControl, _
    ByRef returnedValue)

    returnedValue = VTFormulaFontPresetCount() + 1
End Sub

Public Sub VTWordRibbonGetConversionFontSizeItemLabel( _
    ByVal control As IRibbonControl, _
    ByVal itemIndex As Integer, _
    ByRef returnedValue)

    On Error GoTo LabelFailed
    If itemIndex = 0 Then
        ' "Follow body text"
        returnedValue = VTUnicodeText(36319, 38543, 27491, 25991)
    Else
        returnedValue = VTFormulaFontPresetLabel(itemIndex - 1)
    End If
    Exit Sub

LabelFailed:
    returnedValue = VTUnicodeText(23383, 21495)
    Err.Clear
End Sub

Public Sub VTWordRibbonGetConversionFontSizeSelectedIndex( _
    ByVal control As IRibbonControl, _
    ByRef returnedValue)

    Dim settingPt As Double
    Dim presetIndex As Long

    returnedValue = 0
    On Error GoTo SelectedFinished
    settingPt = VTConversionFontSizeSetting( _
        VTConversionFontSizeControlMode(control))
    If Not VTValidWordFormulaFontSize(settingPt) Then Exit Sub
    presetIndex = VTFormulaFontPresetIndex(settingPt)
    If presetIndex >= 0 Then returnedValue = presetIndex + 1

SelectedFinished:
End Sub

Public Sub VTWordRibbonApplyConversionFontSize( _
    ByVal control As IRibbonControl, _
    ByVal selectedId As String, _
    ByVal selectedIndex As Integer)

    Dim requestedPt As Double
    Dim inlinePt As Double
    Dim displayPt As Double

    On Error GoTo Failed
    If selectedIndex > 0 Then requestedPt = VTFormulaFontPresetSize(selectedIndex - 1)
    ' A size (not "follow body text") resizes every formula of this kind
    ' already in the document first; saving the setting goes through the file
    ' bridge (about 0.3 s) and follows.
    If VTValidWordFormulaFontSize(requestedPt) And Documents.Count > 0 Then
        VTRequireWritableWordDocument
        VTApplyFontSizeToDocumentFormulas _
            VTConversionFontSizeControlMode(control), requestedPt
        VTInvalidateWordImageFontSizeControl
    End If
    inlinePt = VTConversionFontSizeSetting("inline")
    displayPt = VTConversionFontSizeSetting("block")
    If VTConversionFontSizeControlMode(control) = "block" Then
        displayPt = requestedPt
    Else
        inlinePt = requestedPt
    End If
    VTWriteConversionFontSizes inlinePt, displayPt
    If Not VT_WORD_RIBBON Is Nothing Then VT_WORD_RIBBON.InvalidateControl control.Id
    Exit Sub

Failed:
    VTShowError "Word conversion font size", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_SetSelectedImageFormulaFontSize( _
    ByVal enteredText As String)

    Dim formulaShape As InlineShape
    Dim requestedFontSizePt As Double

    If Not VTTryParseWordFormulaFontSize( _
       enteredText, requestedFontSizePt) Then
        Err.Raise vbObjectError + 7496, "VisualTeX", _
            "Enter a numeric image formula font size from 1 to 512 pt, such as 10.5, 12, 14, or 18."
    End If
    Set formulaShape = VTVisualTeXInlineShapeAtSelection(Selection)
    If formulaShape Is Nothing Then
        Err.Raise vbObjectError + 7497, "VisualTeX", _
            "Select one VisualTeX image formula before changing its font size."
    End If
    VTApplyWordImageFormulaFontSize formulaShape, requestedFontSizePt
    formulaShape.Select
End Sub

Public Sub VTBeginWordDoubleClickTraceContext()
    If VT_WORD_DOUBLE_CLICK_TRACE_CONTEXT_DEPTH = 0 Then
        VT_WORD_DOUBLE_CLICK_TRACE_BUFFER = ""
        VT_WORD_DOUBLE_CLICK_TRACE_STARTED_AT = Timer
    End If
    VT_WORD_DOUBLE_CLICK_TRACE_CONTEXT_DEPTH = _
        VT_WORD_DOUBLE_CLICK_TRACE_CONTEXT_DEPTH + 1
End Sub

Public Sub VTEndWordDoubleClickTraceContext()
    Dim traceText As String

    If VT_WORD_DOUBLE_CLICK_TRACE_CONTEXT_DEPTH > 0 Then
        VT_WORD_DOUBLE_CLICK_TRACE_CONTEXT_DEPTH = _
            VT_WORD_DOUBLE_CLICK_TRACE_CONTEXT_DEPTH - 1
    End If
    If VT_WORD_DOUBLE_CLICK_TRACE_CONTEXT_DEPTH <> 0 Or _
       Not VT_WORD_DOUBLE_CLICK_TRACE_ACTIVE Or _
       Len(VT_WORD_DOUBLE_CLICK_TRACE_BUFFER) = 0 Then Exit Sub

    traceText = VT_WORD_DOUBLE_CLICK_TRACE_BUFFER
    VT_WORD_DOUBLE_CLICK_TRACE_BUFFER = ""
    On Error Resume Next
    VTAppendText _
        VTApplicationSupportRoot() & VT_WORD_DOUBLE_CLICK_TRACE_FILE, _
        traceText
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub VTTraceWordDoubleClick( _
    ByVal eventName As String, _
    ByVal selected As Word.Selection, _
    Optional ByVal detail As String = "")

    Dim traceText As String
    Dim eventSinkReady As Boolean
    Dim elapsedMilliseconds As Long

    If Not VT_WORD_DOUBLE_CLICK_TRACE_ACTIVE Then Exit Sub
    If VT_WORD_DOUBLE_CLICK_TRACE_CONTEXT_DEPTH = 0 Then Exit Sub
    If Left$(eventName, 9) = "resolver-" And _
       VT_WORD_DOUBLE_CLICK_TRACE_CONTEXT_DEPTH = 0 Then Exit Sub

    On Error Resume Next
    elapsedMilliseconds = CLng(1000# * VTTimerElapsedSeconds( _
        VT_WORD_DOUBLE_CLICK_TRACE_STARTED_AT, Timer))
    eventSinkReady = Not VT_WORD_EVENT_SINK Is Nothing
    traceText = _
        "event=" & VTDoubleClickTraceValue(eventName) & _
        " time=" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
        " elapsedMs=" & CStr(elapsedMilliseconds) & _
        " revision=" & VT_WORD_SOURCE_REVISION & _
        " wordVersion=" & Application.Version & _
        " documents=" & CStr(Documents.Count) & _
        " internalMutation=" & CStr(VTWordInternalMutationActive()) & _
        " eventSinkReady=" & CStr(eventSinkReady) & _
        " template=" & VTDoubleClickTraceValue(ThisDocument.FullName)
    If selected Is Nothing Then
        traceText = traceText & " selection=nothing"
    Else
        traceText = traceText & _
            " selectionStart=" & CStr(selected.Range.Start) & _
            " selectionEnd=" & CStr(selected.Range.End) & _
            " selectionType=" & CStr(selected.Type) & _
            " selectionInlineShapes=" & CStr(selected.InlineShapes.Count)
    End If
    If Len(detail) > 0 Then
        traceText = traceText & _
            " detail=" & VTDoubleClickTraceValue(detail)
    End If
    traceText = traceText & vbLf
    VT_WORD_DOUBLE_CLICK_TRACE_BUFFER = _
        VT_WORD_DOUBLE_CLICK_TRACE_BUFFER & traceText
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub VisualTeX_WriteWordDoubleClickDiagnosticMarker()
    On Error Resume Next
    VTWriteWordHealth
    VTBeginWordDoubleClickTraceContext
    VTTraceWordDoubleClick "manual-diagnostic-marker", Selection, ""
    VTEndWordDoubleClickTraceContext
    VTShowInformation _
        "VisualTeX Word revision: " & VT_WORD_SOURCE_REVISION & vbCrLf & _
        "Trace: " & VTApplicationSupportRoot() & _
        VT_WORD_DOUBLE_CLICK_TRACE_FILE
    On Error GoTo 0
End Sub

Public Sub VisualTeXPlaceRef()
    ' Native MACROBUTTON entry: Word selects its label on activation. The normal
    ' cross-reference picker supplies reference insertion at the user's caret.
End Sub

' Ribbon: Update Equation Format. Shows what will change, converts every older
' Mac form into the current one (equation format v2) as one Undo step, checks
' the result and undoes it if anything is left or anything else changed.
' The math font step (Windows' Update Equation Format): the chosen OMML font
' becomes the document math font, and every Word equation follows the
' document math font (VTV2MathFontRepairs).
Public Sub VisualTeX_UpdateEquationFormat()
    Dim doc As Document
    Dim numbers As Long
    Dim tags As Long
    Dim baselines As Long
    Dim fonts As Long
    Dim mathFontName As String
    Dim setDocumentFont As Boolean
    Dim inventory As String
    Dim message As String
    Dim recordStarted As Boolean
    Dim mutationStarted As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    VTRequireWritableWordDocument
    Set doc = ActiveDocument
    mathFontName = VTChosenOmmlFontName()
    If Len(mathFontName) > 0 Then
        setDocumentFont = StrComp(VTDocumentMathFontName(doc), _
            mathFontName, vbTextCompare) <> 0
    Else
        mathFontName = VTDocumentMathFontName(doc)
    End If
    numbers = VTV2RestyleNumbers(doc, True)
    tags = VTV2LegacyTagCount(doc)
    baselines = VTV2InlineBaselineRepairs(doc, False)
    fonts = VTV2MathFontRepairs(doc, mathFontName, False)
    If numbers + tags + baselines + fonts = 0 And Not setDocumentFont Then
        VTShowInformation "The equations in this document already use the current format."
        Exit Sub
    End If
    message = "Update this document's equations to the current VisualTeX format?" & vbCr & vbCr
    If numbers > 0 Then message = message & CStr(numbers) & _
        " equation number(s) will be rewritten in the current number format." & vbCr
    If tags > 0 Then message = message & CStr(tags) & _
        " older VisualTeX tag(s) will be removed." & vbCr
    If baselines > 0 Then message = message & CStr(baselines) & _
        " inline image formula(s) will be aligned with the text baseline." & vbCr
    If setDocumentFont Then message = message & _
        "The document math font will be set to " & mathFontName & "." & vbCr
    If fonts > 0 Then message = message & CStr(fonts) & _
        " Word equation(s) will use the document math font (" & mathFontName & _
        "); normal text in equations keeps its font." & vbCr
    message = message & vbCr & "Equation content does not change. Undo reverses the whole update."
    If MsgBox(message, vbOKCancel + vbQuestion, "VisualTeX") <> vbOK Then Exit Sub

    inventory = VTV2EquationInventory(doc)
    Application.UndoRecord.StartCustomRecord "VisualTeX Update Equation Format"
    recordStarted = True
    If setDocumentFont Then doc.OMathFontName = mathFontName
    VTUpgradeWordNumberFormats doc
    VTUpdateCurrentWordNumberFields doc
    VTV2InlineBaselineRepairs doc, True
    If fonts > 0 Then
        VTBeginWordInternalMutation
        mutationStarted = True
        VTV2MathFontRepairs doc, mathFontName, True
        VTEndWordInternalMutation
        mutationStarted = False
    End If
    Application.UndoRecord.EndCustomRecord
    recordStarted = False

    If VTV2RestyleNumbers(doc, True) <> 0 Or VTV2LegacyTagCount(doc) <> 0 Or _
       VTV2InlineBaselineRepairs(doc, False) <> 0 Or _
       VTV2MathFontRepairs(doc, mathFontName, False) <> 0 Or _
       StrComp(VTDocumentMathFontName(doc), mathFontName, vbTextCompare) <> 0 Or _
       VTV2EquationInventory(doc) <> inventory Then
        doc.Undo 1
        Err.Raise vbObjectError + 7613, "VisualTeX", _
            "The equation format update could not be verified and was undone."
    End If
    message = "Updated " & CStr(numbers) & " equation number(s), aligned " & _
        CStr(baselines) & " inline image formula(s) and removed " & _
        CStr(tags) & " older tag(s)."
    If fonts > 0 Or setDocumentFont Then message = message & vbCr & _
        CStr(fonts) & " Word equation(s) changed to the document math font (" & _
        mathFontName & ")."
    VTShowInformation message
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If mutationStarted Then VTEndWordInternalMutation
    If recordStarted Then
        Application.UndoRecord.EndCustomRecord
        doc.Undo 1
    End If
    On Error GoTo 0
    VTShowError "Update Equation Format", failureNumber, failureDescription
End Sub

Public Sub VTWordRibbonGetOmmlFontPressed( _
    ByVal control As IRibbonControl, _
    ByRef returnedValue)

    returnedValue = False
    On Error GoTo Finished
    VTLoadOmmlFontSetting
    returnedValue = (StrComp(VTOmmlControlChoiceKey(control.Id), _
        VTOmmlFontChoiceKey(VT_WORD_OMML_FONT_NAME), vbBinaryCompare) = 0)
Finished:
End Sub

Public Sub VTWordRibbonGetOmmlFontEnabled( _
    ByVal control As IRibbonControl, _
    ByRef returnedValue)

    Dim fontName As String

    returnedValue = True
    On Error GoTo Finished
    fontName = VTOmmlFontNameForControl(control.Id)
    If Len(fontName) > 0 Then returnedValue = VTWordFontInstalled(fontName)
Finished:
End Sub

' The font's name, with "(not installed)" when Word does not list it.
Public Sub VTWordRibbonGetOmmlFontLabel( _
    ByVal control As IRibbonControl, _
    ByRef returnedValue)

    Dim choiceKey As String

    On Error GoTo Finished
    choiceKey = VTOmmlControlChoiceKey(control.Id)
    returnedValue = VTOmmlFontLabelForChoice(choiceKey)
    If Len(choiceKey) > 0 Then
        If Not VTWordFontInstalled(VTOmmlFontNameForChoice(choiceKey)) Then _
            returnedValue = returnedValue & _
                VTUnicodeText(65288, 26410, 23433, 35013, 65289)
    End If
Finished:
End Sub

' A menu item: applies the font to the active document (when it can be
' changed), then saves the choice. "Follow the document" changes nothing.
Public Sub VTWordRibbonSelectOmmlFont( _
    ByVal control As IRibbonControl, _
    ByVal pressed As Boolean)

    Dim fontName As String
    Dim changed As Long
    Dim documentObject As Document

    On Error GoTo Failed
    fontName = VTOmmlFontNameForControl(control.Id)
    If Len(fontName) > 0 Then
        If Not VTWordFontInstalled(fontName) Then
            Err.Raise vbObjectError + 7614, "VisualTeX", _
                VTOmmlFontLabelForChoice(VTOmmlControlChoiceKey(control.Id)) & _
                " is not installed in Word. Install or repair the VisualTeX " & _
                "Office add-ins, then restart Word."
        End If
        If Documents.Count > 0 Then
            Set documentObject = ActiveDocument
            If Not documentObject.ReadOnly And _
               documentObject.ProtectionType = wdNoProtection Then
                changed = VTApplyOmmlFontToDocument(documentObject, fontName)
                Application.StatusBar = "VisualTeX: Word equations use " & _
                    fontName & " (" & CStr(changed) & " equation(s) changed)."
            End If
        End If
    End If
    VTWriteOmmlFontSetting fontName
    VTInvalidateOmmlFontControls
    Exit Sub

Failed:
    VTInvalidateOmmlFontControls
    VTShowError "OMML font", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_NumberDisplayFormulasDocument()
    VTNumberDisplayFormulas True
End Sub

Public Sub VisualTeX_NumberDisplayFormulasSelection()
    VTNumberDisplayFormulas False
End Sub

' Diagnostic (docs/office-mcp/MACOS.md): proves a macro at the end of this module can be run by AppleScript.
Public Sub VisualTeX_AgentPing()
    VTWriteTextAtomic VTSessionRoot() & "/agent-ping.txt", "pong"
End Sub
