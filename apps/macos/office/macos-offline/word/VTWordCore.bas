Attribute VB_Name = "VTWordCore"
Option Explicit
Option Private Module

#Const VT_WORD_DIAGNOSTICS = False

' What the Word modules share: the add-in's names and state, internal mutations, performance traces, document
' variables and identity, the health file, small text and range helpers, and LaTeX to Word's linear format.

' AI agents (VisualTeX_AgentRun in VTWordAdapter, the agent code in VTWordAgent)
Public VT_AGENT_REQUEST As Collection

Public Const VT_WORD_HOST As String = "word"
Private Const VT_WORD_STATUS_FILE As String = "/OfficePluginStatus/word.json"
Public Const VT_WORD_SOURCE_REVISION As String = _
    "word-office-performance-20260801-r130"
Public Const VT_WORD_IDLE_WATCH_SECONDS As Long = 5
Private Const VT_WORD_LATEX_REDRAW_REVISION As String = _
    "word-latex-redraw-20260802-r1"
Private Const VT_WORD_DOCUMENT_IMPORT_REVISION As String = _
    "word-structured-document-import-20260730-r61"
Public Const VT_WORD_EQUATION_NUMBER_FONT_NAME As String = "Cambria Math"
' CoreText glyph-path bounds for "(0)" through "(9)" in the Cambria Math face
' are -0.220703125 em to +0.70556640625 em. Their painted-bbox centre is
' therefore 0.242431640625 em above the Word text baseline. Parentheses own the
' same vertical extrema for chapter/section labels, so the ratio is independent
' of the actual Equation number text.
Public Const VT_WORD_EQUATION_NUMBER_INK_CENTER_ABOVE_BASELINE_RATIO _
    As Double = 0.242431640625#
Public Const VT_WORD_IMAGE_OWNER_BOOKMARK_PREFIX As String = "VT_I_"
Public Const VT_WORD_NATIVE_BOOKMARK_PREFIX As String = "VT_F_"
Public Const VT_WORD_CAPTION_BOOKMARK_PREFIX As String = "VT_C_"
Public Const VT_WORD_NUMBER_BOOKMARK_PREFIX As String = "VT_R_"
Public Const VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX As String = "VT_N_"
Public Const VT_WORD_NUMBERING_MODE_VARIABLE As String = _
    "VT_EquationNumberingMode"
Public Const VT_WORD_NUMBERING_SEPARATOR_VARIABLE As String = _
    "VT_EquationNumberingSeparator"
Public Const VT_WORD_NUMBERING_MODE_SEQUENCE As String = "sequence"
Public Const VT_WORD_NUMBERING_MODE_CHAPTER As String = "chapter"
Public Const VT_WORD_NUMBERING_MODE_SECTION As String = "section"
Public Const VT_WORD_IMAGE_REFERENCE_FONT_SIZE_PT As Double = 14#
Public Const VT_WORD_MAX_FORMULA_FONT_SIZE_PT As Double = 512#
Private Const VT_WORD_INLINE_CONVERSION_FONT_CONTROL_ID As String = _
    "VisualTeX.Mac.Word.InlineConversionFontSize"
Public Const VT_WORD_IMAGE_MACRO_SCHEMA_VARIABLE As String = _
    "VT_ImageMacroButtonSchema"
Public Const VT_WORD_IMAGE_MACRO_SCHEMA_VERSION As String = "10"
Public Const VT_WORD_IMAGE_SEQUENCE_NAME As String = "VisualTeXEquation"
Public Const VT_WORD_REFERENCE_KIND_NATIVE As String = "omml"
Public Const VT_WORD_REFERENCE_KIND_IMAGE As String = "image"
Public Const VT_WORD_IMAGE_EDIT_DEBOUNCE_SECONDS As Double = 0.75
Private Const VT_WORD_OMML_LATEX_RESULT_FILE As String = _
    "/omml-latex-result.txt"
Public Const VT_WORD_LATEX_REDRAW_MAX_BYTES As Long = 5242880
Public Const VT_WORD_LATEX_REDRAW_MAX_FORMULAS As Long = 1000
Public Const VT_WORD_TRACE_ENABLED As Boolean = False
Private Const VT_WORD_PERFORMANCE_TRACE_FILE As String = _
    "/word-vba-performance.txt"
Private Const VT_WORD_PERFORMANCE_TRACE_SENTINEL As String = _
    "/word-vba-performance.enabled"
Public Const VT_WORD_OMML_CHUNK_SIZE As Long = VT_WORD_PAYLOAD_CHUNK_SIZE
Public Const VT_WORD_OMML_MAX_CHUNKS As Long = VT_WORD_PAYLOAD_MAX_CHUNKS
Public VT_WORD_RIBBON As IRibbonUI
Public VT_WORD_INTERNAL_MUTATION_DEPTH As Long
Public VT_WORD_NUMBER_UPDATE_RUNNING As Boolean
' The numbered image paragraph style while a batch numbers images: prepared
' once, because changing a style makes Word revisit every paragraph using it.
Public VT_WORD_DISPLAY_NUMBERING_IMAGE_STYLE As Style
' The text width of a one-section document while a batch numbers images
' (VTEquationLayoutWidth asks the section's page setup for every formula).
Public VT_WORD_DISPLAY_NUMBERING_LAYOUT_WIDTH As Single
Private VT_WORD_MUTATION_SMART_CUT_PASTE As Boolean
Public VT_WORD_V2_IMPORT_CACHE_ACTIVE As Boolean
Public VT_WORD_V2_IMPORT_ENTRIES As Collection
Public VT_WORD_V2_IMPORT_HEADINGS As Collection
Public VT_WORD_V2_IMPORT_HEADINGS_END As Long
Public VT_WORD_V2_IMPORT_HEADINGS_LEVEL As Long
Public VT_WORD_DOUBLE_CLICK_TRACE_ACTIVE As Boolean
Public VT_WORD_PERFORMANCE_ACTIVE As Boolean
Private VT_WORD_PERFORMANCE_STARTED_AT As Single
Public VT_WORD_PERFORMANCE_BUFFER As String
Private VT_WORD_PERFORMANCE_OPERATION As String
Public VT_WORD_IMAGE_SIZE_WATCH_SCHEDULED As Boolean
' True while the caret, the idle watcher or document open re-aligns image
' formulas: Equation numbers then keep the size, font and colour they have,
' and only creating, resizing or renumbering a formula sets them.
Public VT_WORD_KEEP_NUMBER_FORMAT As Boolean
Public VT_WORD_IMAGE_WATCH_LAST_DELAY As Long
Public VT_WORD_IDLE_TASK_DUE_AT As Date
Public VT_WORD_IMAGE_SIZE_WATCH_RUNNING As Boolean
Public VT_WORD_IMAGE_CROP_CACHE As Collection
Public VT_WORD_ORPHAN_WATCH_SCHEDULED As Boolean
Public VT_WORD_ORPHAN_WATCH_RUNNING As Boolean
Public VT_WORD_OMML_FONT_NAME As String

Public Function VTWordVbeBuildTemplateActive() As Boolean
    Dim projectDocumentName As String
    Dim projectDocumentPath As String

    On Error Resume Next
    projectDocumentName = ThisDocument.Name
    projectDocumentPath = ThisDocument.FullName
    On Error GoTo 0
    VTWordVbeBuildTemplateActive = _
        InStr(1, projectDocumentName, "VisualTeXWordBuild", vbTextCompare) > 0 Or _
        InStr(1, projectDocumentPath, "/VisualTeX/Scratch/VisualTeXWordBuild", _
            vbTextCompare) > 0
End Function

Public Sub VTBeginWordInternalMutation()
#If VT_WORD_DIAGNOSTICS Then
    VT_WORD_REGRESSION_MUTATION_COUNT = VT_WORD_REGRESSION_MUTATION_COUNT + 1
#End If
    If VT_WORD_INTERNAL_MUTATION_DEPTH = 0 Then
        ' Range.Delete also applies Word's smart spacing. With "text $x$ text"
        ' it can delete a neighbouring space outside the requested range,
        ' invalidating the offsets of both this and earlier replacements.
        ' Every programmatic formula transaction owns exact character ranges;
        ' preserve them through delete, transfer, and rollback. Nested calls
        ' must retain the outer transaction's original user preference.
        VT_WORD_MUTATION_SMART_CUT_PASTE = Options.SmartCutPaste
        Options.SmartCutPaste = False
    End If
    VT_WORD_INTERNAL_MUTATION_DEPTH = _
        VT_WORD_INTERNAL_MUTATION_DEPTH + 1
End Sub

Public Sub VTEndWordInternalMutation()
    If VT_WORD_INTERNAL_MUTATION_DEPTH > 0 Then
        VT_WORD_INTERNAL_MUTATION_DEPTH = _
            VT_WORD_INTERNAL_MUTATION_DEPTH - 1
        If VT_WORD_INTERNAL_MUTATION_DEPTH = 0 Then
            Options.SmartCutPaste = VT_WORD_MUTATION_SMART_CUT_PASTE
        End If
    End If
End Sub

Public Sub VTCollectionAddUniqueText( _
    ByVal values As Collection, _
    ByVal value As String)

    If values Is Nothing Or Len(value) = 0 Then Exit Sub
    On Error Resume Next
    values.Add value, LCase$(value)
    On Error GoTo 0
End Sub

Public Function VTCollectionContainsText( _
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

Public Function VTTryGetBookmarkRangeIncludingHidden( _
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

Public Sub VTReplaceBookmarkRangeIncludingHidden( _
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

Public Function VTFirstPositiveIntegerInText( _
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

Public Function VTLastIntegerInText( _
    ByVal sourceText As String) As Long

    Dim characterIndex As Long
    Dim currentCharacter As String
    Dim digitText As String
    Dim lastValue As Long

    For characterIndex = 1 To Len(sourceText)
        currentCharacter = Mid$(sourceText, characterIndex, 1)
        If currentCharacter >= "0" And currentCharacter <= "9" Then
            digitText = digitText & currentCharacter
        ElseIf Len(digitText) > 0 Then
            lastValue = CLng(Val(digitText))
            digitText = ""
        End If
    Next characterIndex
    If Len(digitText) > 0 Then lastValue = CLng(Val(digitText))
    VTLastIntegerInText = lastValue
End Function

Public Function VTVariantArrayCount(ByVal values As Variant) As Long
    On Error GoTo NoValues
    If IsArray(values) Then
        VTVariantArrayCount = UBound(values) - LBound(values) + 1
    End If
    Exit Function
NoValues:
    VTVariantArrayCount = 0
End Function

Public Function VTNormalizedWordFieldCode(ByVal fieldCode As String) As String
    fieldCode = Replace$(fieldCode, vbTab, " ")
    fieldCode = Replace$(fieldCode, ChrW(160), " ")
    fieldCode = Trim$(fieldCode)
    Do While InStr(1, fieldCode, "  ", vbBinaryCompare) > 0
        fieldCode = Replace$(fieldCode, "  ", " ")
    Loop
    VTNormalizedWordFieldCode = fieldCode
End Function

Public Function VTTryVisualTeXMetadataShapeAtSelection( _
    ByVal selected As Selection, _
    ByRef resolvedShape As InlineShape) As Boolean

    Dim candidate As InlineShape
    Dim candidateField As Field
    Dim probeRange As Range
    Dim match As InlineShape
    Dim matchCount As Long
    Dim fieldStart As Long
    Dim fieldEnd As Long

    On Error GoTo InvalidSelection
    Set resolvedShape = Nothing
    If selected Is Nothing Then Exit Function

    If selected.InlineShapes.Count = 1 Then
        Set candidate = selected.InlineShapes(1)
        If VTInlineShapeHasVisualTeXMetadata(candidate) Then
            Set resolvedShape = candidate
            VTTryVisualTeXMetadataShapeAtSelection = True
            Exit Function
        End If
    End If

    For Each candidateField In selected.Range.Paragraphs(1).Range.Fields
        If VTIsVisualTeXImageMacroButtonField(candidateField) And _
           VTMacroButtonResultHasOnlyImage(candidateField) Then
            fieldStart = candidateField.Code.Start - 1
            fieldEnd = candidateField.Result.End + 1
            If selected.Range.Start <= fieldEnd + 1 And _
               selected.Range.End >= fieldStart - 1 Then
                Set candidate = candidateField.Result.InlineShapes(1)
                If VTInlineShapeHasVisualTeXMetadata(candidate) Then
                    matchCount = matchCount + 1
                    Set match = candidate
                End If
            End If
        End If
    Next candidateField
    If matchCount = 1 Then
        Set resolvedShape = match
        VTTryVisualTeXMetadataShapeAtSelection = True
        Exit Function
    End If

    matchCount = 0
    Set match = Nothing
    Set probeRange = selected.Range.Duplicate
    probeRange.MoveStart Unit:=wdCharacter, Count:=-1
    probeRange.MoveEnd Unit:=wdCharacter, Count:=1
    For Each candidate In probeRange.InlineShapes
        If VTInlineShapeHasVisualTeXMetadata(candidate) Then
            matchCount = matchCount + 1
            Set match = candidate
        End If
    Next candidate
    If matchCount = 1 Then
        Set resolvedShape = match
        VTTryVisualTeXMetadataShapeAtSelection = True
    End If
    Exit Function

InvalidSelection:
    Set resolvedShape = Nothing
    VTTryVisualTeXMetadataShapeAtSelection = False
End Function

Public Function VTTimerElapsedSeconds( _
    ByVal startedAt As Single, _
    ByVal finishedAt As Single) As Double

    If finishedAt < startedAt Then finishedAt = finishedAt + 86400!
    VTTimerElapsedSeconds = CDbl(finishedAt) - CDbl(startedAt)
End Function

Public Sub VTWordPerformanceStart()
    VT_WORD_PERFORMANCE_STARTED_AT = Timer
    VT_WORD_PERFORMANCE_ACTIVE = False
    VT_WORD_PERFORMANCE_BUFFER = ""
    VT_WORD_PERFORMANCE_OPERATION = ""
End Sub

Public Sub VTWordPerformanceStartLocal( _
    ByVal operationName As String)

    VTWordPerformanceStart
    If Not VTPathFileExists( _
       VTApplicationSupportRoot() & _
       VT_WORD_PERFORMANCE_TRACE_SENTINEL) Then Exit Sub
    VT_WORD_PERFORMANCE_ACTIVE = True
    VT_WORD_PERFORMANCE_OPERATION = operationName
    VTWordPerformanceMark "local-start"
End Sub

Public Sub VTWordPerformanceEnable( _
    ByVal dispatch As Object)

    If dispatch Is Nothing Then Exit Sub
    VT_WORD_PERFORMANCE_ACTIVE = _
        (VTDispatchOptional(dispatch, "performanceTrace") = "1")
    If Not VT_WORD_PERFORMANCE_ACTIVE Then Exit Sub
    VT_WORD_PERFORMANCE_OPERATION = _
        VTDispatchOptional(dispatch, "operation")
    If Len(VT_WORD_PERFORMANCE_OPERATION) = 0 Then
        VT_WORD_PERFORMANCE_OPERATION = "formula"
    End If
    VTWordPerformanceMark "dispatch-read"
End Sub

Public Sub VTWordPerformanceMark(ByVal stageName As String)
    Dim elapsedMilliseconds As Long

    If Not VT_WORD_PERFORMANCE_ACTIVE Then Exit Sub
    elapsedMilliseconds = CLng(1000# * VTTimerElapsedSeconds( _
        VT_WORD_PERFORMANCE_STARTED_AT, Timer))
    VT_WORD_PERFORMANCE_BUFFER = VT_WORD_PERFORMANCE_BUFFER & _
        "stage=" & stageName & _
        ";elapsedMs=" & CStr(elapsedMilliseconds) & vbLf
End Sub

Public Sub VTWordPerformanceNote( _
    ByVal keyName As String, _
    ByVal valueText As String)

    If Not VT_WORD_PERFORMANCE_ACTIVE Then Exit Sub
    VT_WORD_PERFORMANCE_BUFFER = VT_WORD_PERFORMANCE_BUFFER & _
        keyName & "=" & Replace$(Replace$( _
            valueText, vbCr, " "), vbLf, " ") & vbLf
End Sub

Public Sub VTWordPerformanceFlush( _
    ByVal sessionId As String, _
    Optional ByVal errorNumber As Long = 0, _
    Optional ByVal errorDescription As String = "", _
    Optional ByVal traceFile As String = VT_WORD_PERFORMANCE_TRACE_FILE)

    Dim traceText As String

    If Not VT_WORD_PERFORMANCE_ACTIVE Or Len(sessionId) = 0 Then Exit Sub
    On Error Resume Next
    traceText = _
        "schema=visualtex-word-vba-performance-v1" & vbLf & _
        "operation=" & VT_WORD_PERFORMANCE_OPERATION & vbLf & _
        VT_WORD_PERFORMANCE_BUFFER & _
        "errorNumber=" & CStr(errorNumber) & vbLf & _
        "errorDescription=" & Replace$(Replace$( _
            errorDescription, vbCr, " "), vbLf, " ") & vbLf
    VTWriteTextAtomic _
        VTSessionDirectory(sessionId) & traceFile, _
        traceText
    Err.Clear
    On Error GoTo 0
    VT_WORD_PERFORMANCE_ACTIVE = False
    VT_WORD_PERFORMANCE_BUFFER = ""
    VT_WORD_PERFORMANCE_OPERATION = ""
End Sub

' True when a range holds nothing but spacing: no text, field, image or equation.
Public Function VTRangeIsBlankSpacing(ByVal candidateRange As Range) As Boolean
    If candidateRange.End <= candidateRange.Start Then
        VTRangeIsBlankSpacing = True
        Exit Function
    End If
    If candidateRange.InlineShapes.Count > 0 Or _
       candidateRange.OMaths.Count > 0 Or _
       candidateRange.Fields.Count > 0 Then Exit Function
    VTRangeIsBlankSpacing = Not VTWordRangeHasMeaningfulText(candidateRange)
End Function

Public Function VTXmlTextContent(ByVal xmlFragment As String) As String
    Dim tagStart As Long
    Dim tagEnd As Long

    Do
        tagStart = InStr(1, xmlFragment, "<", vbBinaryCompare)
        If tagStart = 0 Then Exit Do
        tagEnd = InStr(tagStart, xmlFragment, ">", vbBinaryCompare)
        If tagEnd = 0 Then Exit Function
        xmlFragment = Left$(xmlFragment, tagStart - 1) & _
            Mid$(xmlFragment, tagEnd + 1)
    Loop
    xmlFragment = Replace$(xmlFragment, vbCr, "")
    xmlFragment = Replace$(xmlFragment, vbLf, "")
    xmlFragment = Replace$(xmlFragment, vbTab, "")
    xmlFragment = Replace$(xmlFragment, " ", "")
    VTXmlTextContent = xmlFragment
End Function

Public Function VTXmlElementStartBefore( _
    ByVal xmlText As String, _
    ByVal elementName As String, _
    ByVal beforePosition As Long) As Long

    Dim candidate As Long
    Dim boundary As String
    Dim prefix As String

    prefix = "<" & elementName
    If beforePosition <= 1 Or Len(prefix) = 1 Then Exit Function
    candidate = InStrRev( _
        Left$(xmlText, beforePosition - 1), prefix, -1, vbBinaryCompare)
    Do While candidate > 0
        boundary = Mid$(xmlText, candidate + Len(prefix), 1)
        If boundary = ">" Or boundary = " " Or boundary = vbTab Or _
           boundary = vbCr Or boundary = vbLf Then
            VTXmlElementStartBefore = candidate
            Exit Function
        End If
        If candidate <= 1 Then Exit Do
        candidate = InStrRev( _
            Left$(xmlText, candidate - 1), prefix, -1, vbBinaryCompare)
    Loop
End Function

Public Function VTWordRangeHasMeaningfulText( _
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

Public Function VTTryGetDocumentVariable( _
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

Public Sub VTSetDocumentVariable( _
    ByVal documentObject As Document, _
    ByVal variableName As String, _
    ByVal variableValue As String)

    Dim ignored As String
    If VTTryGetDocumentVariable(documentObject, variableName, ignored) Then
        If StrComp(ignored, variableValue, vbBinaryCompare) <> 0 Then _
            documentObject.Variables(variableName).Value = variableValue
    Else
        documentObject.Variables.Add Name:=variableName, Value:=variableValue
    End If
End Sub

Public Sub VTDeleteDocumentVariable(ByVal documentObject As Document, ByVal variableName As String)
    On Error Resume Next
    documentObject.Variables(variableName).Delete
    Err.Clear
    On Error GoTo 0
End Sub

' Range.Information(wdWithInTable) makes Word for Mac lay the document out
' first, which costs far more than the answer after an edit in a long
' document. A Range without any Table is outside every table, so only a Range
' that touches a table asks Word.
Public Function VTRangeIsWithinTable(ByVal targetRange As Range) As Boolean
    If targetRange Is Nothing Then Exit Function
    If targetRange.Tables.Count = 0 Then Exit Function
    VTRangeIsWithinTable = targetRange.Information(wdWithInTable)
End Function

Public Sub VTScheduleWordIdleTasks(ByVal delaySeconds As Long)
    Dim requestedTime As Date
    requestedTime = Now + TimeSerial(0, 0, delaySeconds)
    ' Word owns one OnTime slot. A quiet five-second watcher must not replace
    ' the earlier one-second startup/prewarm callback, or postpone pending work
    ' on each selection event. Keep the earliest requested deadline.
    If VT_WORD_IDLE_TASK_DUE_AT > Now And _
       VT_WORD_IDLE_TASK_DUE_AT <= requestedTime Then Exit Sub
    Application.OnTime When:=requestedTime, name:="VisualTeX_RunWordIdleTasks"
    VT_WORD_IDLE_TASK_DUE_AT = requestedTime
End Sub

Public Function VTLaTeXToWordLinear(ByVal latex As String) As String
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

Public Sub VTRequireWritableWordDocument()
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

Public Function VTWordDocumentIdentityForDocument( _
    ByVal documentObject As Document) As String

    On Error Resume Next
    VTWordDocumentIdentityForDocument = documentObject.FullName
    If Err.Number <> 0 Or Len(VTWordDocumentIdentityForDocument) = 0 Then
        Err.Clear
        VTWordDocumentIdentityForDocument = documentObject.Name
    End If
    On Error GoTo 0
    VTWordDocumentIdentityForDocument = _
        VTBoundedIdentity(VTWordDocumentIdentityForDocument)
End Function

Public Function VTWordDocumentIdentity() As String
    If Documents.Count = 0 Then Exit Function
    VTWordDocumentIdentity = _
        VTWordDocumentIdentityForDocument(ActiveDocument)
End Function

Private Function VTWordDocumentIdentityLeaf( _
    ByVal identity As String) As String

    Dim separator As Long

    identity = Trim$(identity)
    If Len(identity) = 0 Then Exit Function
    separator = InStrRev(identity, "/")
    If separator = 0 Then separator = InStrRev(identity, ":")
    If separator > 0 And separator < Len(identity) Then
        VTWordDocumentIdentityLeaf = Mid$(identity, separator + 1)
    Else
        VTWordDocumentIdentityLeaf = identity
    End If
End Function

Public Function VTFindOpenWordDocumentByIdentity( _
    ByVal expectedIdentity As String) As Document

    Dim candidate As Document
    Dim fallback As Document
    Dim expectedName As String
    Dim fallbackCount As Long

    If Len(expectedIdentity) = 0 Then Exit Function
    expectedName = VTWordDocumentIdentityLeaf(expectedIdentity)
    For Each candidate In Documents
        If VTWordDocumentIdentityForDocument(candidate) = expectedIdentity Then
            Set VTFindOpenWordDocumentByIdentity = candidate
            Exit Function
        End If
        If Len(expectedName) > 0 And _
           StrComp(candidate.Name, expectedName, vbBinaryCompare) = 0 Then
            fallbackCount = fallbackCount + 1
            Set fallback = candidate
        End If
    Next candidate
    If fallbackCount = 1 Then
        Set VTFindOpenWordDocumentByIdentity = fallback
    End If
End Function

Public Sub VTTraceWordSession( _
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

Public Sub VTWriteWordFailureTrace( _
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

Public Sub VTRefreshWordHealthQuietly()
    On Error Resume Next
    VTWriteWordHealth
    On Error GoTo 0
End Sub

Public Sub VTWriteWordHealth()
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
