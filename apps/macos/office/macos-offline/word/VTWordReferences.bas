Attribute VB_Name = "VTWordReferences"
Option Explicit
Option Private Module

' Cross-references to equation numbers: reference targets and fields, their formatting and repair.

Public Function VTReferenceTargetBookmarkName( _
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

Public Function VTReferenceResultMatchesEquationNumber( _
    ByVal referenceResultText As String, _
    ByVal equationNumberText As String) As Boolean

    Dim comparableResult As String
    Dim comparableNumber As String
    Dim prefixText As String
    Dim prefixCharacter As String

    comparableResult = VTComparableEquationNumberText( _
        Trim$(referenceResultText))
    comparableNumber = VTComparableEquationNumberText( _
        Trim$(equationNumberText))
    If Len(comparableNumber) = 0 Then Exit Function
    If comparableResult = comparableNumber Or _
       comparableResult = "(" & comparableNumber & ")" Then
        VTReferenceResultMatchesEquationNumber = True
        Exit Function
    End If
    If Len(comparableResult) <= Len(comparableNumber) Then Exit Function
    If Right$(comparableResult, Len(comparableNumber)) <> _
       comparableNumber Then Exit Function

    prefixText = Left$(comparableResult, _
        Len(comparableResult) - Len(comparableNumber))
    prefixText = RTrim$(prefixText)
    If Len(prefixText) = 0 Then
        VTReferenceResultMatchesEquationNumber = True
        Exit Function
    End If
    prefixCharacter = Right$(prefixText, 1)
    If InStr(1, "0123456789.-", prefixCharacter, vbBinaryCompare) = 0 Then
        VTReferenceResultMatchesEquationNumber = True
    End If
End Function

Private Function VTFormulaIdForMisanchoredPrivateReference( _
    ByVal documentObject As Document, _
    ByVal targetRange As Range, _
    ByVal formulaIds As Variant) As String

    Dim formulaId As String
    Dim formulaRange As Range
    Dim formulaParagraph As Range
    Dim itemIndex As Long

    If documentObject Is Nothing Or targetRange Is Nothing Or _
       Not IsArray(formulaIds) Then Exit Function

    ' Word for Mac can anchor a new private _Ref at the first position of the
    ' visible formula paragraph even though its native caption is the following
    ' framed helper. Resolve only that exact numbered-formula containment; an
    ' arbitrary private Bookmark elsewhere in the document is never rewritten.
    For itemIndex = LBound(formulaIds) To UBound(formulaIds)
        formulaId = CStr(formulaIds(itemIndex))
        Set formulaRange = VTNumberedFormulaRangeForId( _
            documentObject, formulaId)
        If Not formulaRange Is Nothing Then
            Set formulaParagraph = _
                VTWordParagraphContainingFormula(formulaRange)
            If Not formulaParagraph Is Nothing Then
                If targetRange.Start >= formulaParagraph.Start And _
                   targetRange.Start < formulaParagraph.End And _
                   targetRange.End <= formulaParagraph.End Then
                    VTFormulaIdForMisanchoredPrivateReference = formulaId
                    Exit Function
                End If
            End If
        End If
    Next itemIndex
End Function

Public Function VTFormulaIdForReferenceTarget( _
    ByVal documentObject As Document, _
    ByVal targetBookmarkName As String, _
    ByRef targetKind As String, _
    Optional ByVal referenceResultText As String = "", _
    Optional ByVal knownFormulaIds As Variant) As String

    Dim formulaIds As Variant
    Dim formulaId As String
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim captionBookmarkName As String
    Dim targetRange As Range
    Dim candidateRange As Range
    Dim itemIndex As Long
    Dim referenceOrdinal As Long
    Dim matchingFormulaId As String
    Dim matchingFormulaCount As Long
    Dim expectedNumberText As String

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
    If IsMissing(knownFormulaIds) Then
        formulaIds = VTValidNumberedFormulaIds(documentObject)
    Else
        formulaIds = knownFormulaIds
    End If
    If Not IsArray(formulaIds) Then Exit Function
    If VTTryGetBookmarkRangeIncludingHidden( _
       documentObject, targetBookmarkName, targetRange) Then
        ' A framed caption can make Word for Mac expose a newly generated private
        ' Bookmark as Start=frame-anchor, End=formula-start. Normalize only that
        ' reversed _Ref view for identity matching; replacement still uses the
        ' original Bookmark name and therefore preserves the native REF field.
        If Left$(targetBookmarkName, 4) = "_Ref" And _
           targetRange.Start > targetRange.End Then
            Set targetRange = documentObject.Range( _
                Start:=targetRange.End, End:=targetRange.Start)
        End If
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

    If Left$(targetBookmarkName, 4) = "_Ref" And _
       Not targetRange Is Nothing Then
        formulaId = VTFormulaIdForMisanchoredPrivateReference( _
            documentObject, targetRange, formulaIds)
        If Len(formulaId) > 0 Then
            targetKind = "N"
            VTFormulaIdForReferenceTarget = formulaId
            Exit Function
        End If
    End If

    ' Word for Mac can create a private _Ref Bookmark whose Range is not
    ' stably nested inside the native SEQ/caption Range. In chapter/section
    ' modes the first integer in a visible result such as 2-2 is the heading
    ' prefix, not the document-order Equation ordinal. Match the complete
    ' formatted number first so a private _Ref remains bound to the same formula
    ' when a local ordinal changes or that formula gains a native \\r 1 reset.
    For itemIndex = LBound(formulaIds) To UBound(formulaIds)
        formulaId = CStr(formulaIds(itemIndex))
        expectedNumberText = VTEquationNumberTextForFormula( _
            documentObject, formulaId)
        If VTReferenceResultMatchesEquationNumber( _
           referenceResultText, expectedNumberText) Then
            matchingFormulaCount = matchingFormulaCount + 1
            matchingFormulaId = formulaId
        End If
    Next itemIndex
    If matchingFormulaCount = 1 Then
        targetKind = "N"
        VTFormulaIdForReferenceTarget = matchingFormulaId
        Exit Function
    End If

    ' Plain sequential numbering has no chapter prefix, so the legacy ordinal
    ' fallback remains deterministic when Word exposes neither an overlapping
    ' private Bookmark Range nor a uniquely comparable result string.
    If VTEquationNumberingMode(documentObject) = _
       VT_WORD_NUMBERING_MODE_SEQUENCE Then
        referenceOrdinal = VTFirstPositiveIntegerInText(referenceResultText)
        If referenceOrdinal >= 1 And _
           referenceOrdinal <= VTVariantArrayCount(formulaIds) Then
            targetKind = "N"
            VTFormulaIdForReferenceTarget = _
                CStr(formulaIds(LBound(formulaIds) + referenceOrdinal - 1))
        End If
    End If
End Function

Public Function VTCaptureBodyEquationReferenceBindings( _
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

Public Sub VTRestoreBodyEquationReferenceBindings( _
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

Public Sub VTReplaceDeletedEquationBodyReferences( _
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

Public Sub VTNormalizeSelectedEquationReferences()
    Dim scanRange As Range

    If Documents.Count = 0 Then Exit Sub
    Set scanRange = Selection.Range.Duplicate
    If scanRange.Start = scanRange.End Then
        Set scanRange = scanRange.Paragraphs(1).Range.Duplicate
    End If
    If scanRange.Fields.Count = 0 Then Exit Sub

    ' Repair manual F9/context-menu updates only where the user is working.
    ' A collapsed caret scans one paragraph; an explicit multi-paragraph
    ' selection scans that selection. Idle typing no longer enumerates every
    ' REF field in a document once per second.
    VTNormalizeBodyEquationReferenceVisibilityInRange _
        scanRange.Document, scanRange
End Sub

Public Function VTEquationNumberCrossReferenceItems( _
    ByVal documentObject As Document, _
    Optional ByVal referenceKind As String = "") As Variant
    VTEquationNumberCrossReferenceItems = VTNumberTargetItems( _
        VTCurrentNumberTargets(documentObject, referenceKind))
End Function

' The labels of a list of numbers ("(3)"), or Empty for none.
Public Function VTNumberTargetItems(ByVal targets As Collection) As Variant
    Dim index As Long
    Dim items() As String
    If targets Is Nothing Then Exit Function
    If targets.Count = 0 Then Exit Function
    ReDim items(1 To targets.Count)
    For index = 1 To targets.Count
        items(index) = "(" & Trim$(targets(index).Text) & ")"
    Next index
    VTNumberTargetItems = items
End Function

Public Function VTResolveEquationReferenceInsertionRange( _
    ByVal requestedRange As Range) As Range

    Dim documentObject As Document
    Dim targetParagraph As Range
    Dim candidateField As Field
    Dim sequenceBookmarkName As String
    Dim formulaId As String
    Dim formulaRange As Range
    Dim recoveredRange As Range

    If requestedRange Is Nothing Then Exit Function
    Set documentObject = requestedRange.Document
    Set targetParagraph = requestedRange.Paragraphs(1).Range.Duplicate

    If Not VTParagraphHasNumberHelperGeometry(targetParagraph) Then
        Set VTResolveEquationReferenceInsertionRange = requestedRange.Duplicate
        Exit Function
    End If

    ' The compact helper paragraph is an internal implementation detail and is
    ' not directly reachable through normal document editing. If Word leaves the
    ' Selection there after a numbered display mutation, recover locally through
    ' that helper's own VT_N_ identity and move to the formula's plain
    ' continuation paragraph. This is O(1) with respect to healthy tail inserts
    ' and does not prune/reconcile the document.
    For Each candidateField In targetParagraph.Fields
        If VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            sequenceBookmarkName = VTSequenceBookmarkNameForField( _
                documentObject, candidateField)
            formulaId = VTFormulaIdFromSequenceBookmarkName( _
                sequenceBookmarkName)
            If Len(formulaId) > 0 Then Exit For
        End If
    Next candidateField

    If Len(formulaId) = 0 Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The Equation cross-reference insertion point is inside a damaged " & _
            "numbering helper with no live formula identity. Run Update Numbers first."
    End If

    Set formulaRange = VTNumberedFormulaRangeForId(documentObject, formulaId)
    If formulaRange Is Nothing Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "The Equation cross-reference insertion point lost its numbered formula."
    End If
    Set recoveredRange = VTContinuationRangeAfterDisplayFormula( _
        formulaRange, formulaId)
    If recoveredRange Is Nothing Then
        Err.Raise vbObjectError + 7547, "VisualTeX", _
            "Word could not recover a normal insertion point after the Equation."
    End If
    Set VTResolveEquationReferenceInsertionRange = recoveredRange.Duplicate
End Function

' knownTargets: the scan of the numbers (VTCurrentNumberTargets of this kind) the caller has just made, when nothing has changed in
' the document since; without it the numbers are scanned here.
Public Function VTInsertEquationNumberReferenceAtRange( _
    ByVal targetRange As Range, _
    ByVal itemIndex As Long, _
    Optional ByVal documentAlreadyReconciled As Boolean = False, _
    Optional ByVal referenceKind As String = "", _
    Optional ByVal knownTargets As Collection = Nothing) As Range
    Dim targets As Collection
    Dim doc As Document
    Dim source As Range
    Dim insertion As Range
    Dim reference As Field
    Dim bookmarkName As String
    Dim start As Long
    If targetRange Is Nothing Then Err.Raise vbObjectError + 7547
    Set doc = targetRange.Document
    If knownTargets Is Nothing Then
        Set targets = VTCurrentNumberTargets(doc, referenceKind)
    Else
        Set targets = knownTargets
    End If
    VTWordPerformanceMark "xref-targets-scanned"
    VTWordPerformanceNote "xref-targets", CStr(targets.Count)
    If itemIndex < 1 Or itemIndex > targets.Count Then Err.Raise vbObjectError + 7547, , "Number target index=" & CStr(itemIndex) & "; count=" & CStr(targets.Count)
    Set source = targets(itemIndex)
    bookmarkName = "_Ref" & Replace$(VTNewUuidV4(), "-", "")
    doc.Bookmarks.Add Name:=bookmarkName, Range:=source
    VTWordPerformanceMark "xref-bookmark-added"
    Set insertion = VTResolveEquationReferenceInsertionRange(targetRange)
    VTWordPerformanceMark "xref-range-resolved"
    start = insertion.Start
    insertion.Text = "()"
    Set insertion = doc.Range(start + 1, start + 1)
    VTWordPerformanceMark "xref-text-inserted"
    Set reference = doc.Fields.Add(Range:=insertion, Type:=wdFieldRef, _
        Text:=bookmarkName & " \h", PreserveFormatting:=False)
    VTWordPerformanceMark "xref-field-added"
    reference.Update
    VTWordPerformanceMark "xref-field-updated"
    Set insertion = doc.Range(start, VTEquationFieldEnd(reference) + 1)
    VTFormatBodyEquationReference doc, reference, insertion
    VTWordPerformanceMark "xref-formatted"
    Set VTInsertEquationNumberReferenceAtRange = insertion

End Function

Public Function VTEquationCrossReferenceText( _
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

Private Function VTParenthesizedEquationReferenceFieldText( _
    ByVal targetBookmarkName As String) As String

    If Len(Trim$(targetBookmarkName)) = 0 Then
        Err.Raise vbObjectError + 7546, "VisualTeX", _
            "The Equation reference Bookmark name is missing."
    End If
    ' Body cross-references remain hyperlinks so a reader can jump back to the
    ' numbered formula. The visible number printed beside a formula must not use
    ' this switch: Word for Mac rebuilds a clicked hyperlink REF result from its
    ' source Bookmark and silently drops the image-number baseline correction.
    VTParenthesizedEquationReferenceFieldText = _
        targetBookmarkName & " \h"
End Function

Public Function VTVisibleEquationReferenceFieldText( _
    ByVal targetBookmarkName As String) As String

    If Len(Trim$(targetBookmarkName)) = 0 Then
        Err.Raise vbObjectError + 7546, "VisualTeX", _
            "The visible Equation reference Bookmark name is missing."
    End If
    VTVisibleEquationReferenceFieldText = targetBookmarkName
End Function

Public Function VTImageEquationReferenceField( _
    ByVal formulaRange As Range, _
    ByVal formulaId As String) As Field

    Dim exactRange As Range
    Dim paragraphRange As Range
    Dim candidateField As Field
    Dim targetBookmarkName As String
    Dim candidateTargetName As String
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim matchCount As Long
    Dim match As Field

    Set VTImageEquationReferenceField = VTImagePlaceRef(formulaRange)
    If Not VTImageEquationReferenceField Is Nothing Then Exit Function
    If formulaRange Is Nothing Then Exit Function
    If formulaRange.InlineShapes.Count <> 1 Or _
       formulaRange.OMaths.Count <> 0 Or _
       Not VTIsCanonicalUuid(formulaId) Then Exit Function
    Set exactRange = formulaRange.InlineShapes(1).Range.Duplicate
    Set paragraphRange = exactRange.Paragraphs(1).Range.Duplicate
    targetBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    For Each candidateField In paragraphRange.Fields
        If candidateField.Type = wdFieldRef Then
            fieldStart = VTEquationFieldStart(candidateField)
            fieldEnd = VTEquationFieldEnd(candidateField)
            If fieldStart > exactRange.End And _
               fieldEnd < paragraphRange.End Then
                candidateTargetName = VTReferenceTargetBookmarkName( _
                    candidateField.Code.Text)
                If StrComp(candidateTargetName, targetBookmarkName, _
                   vbTextCompare) = 0 Then
                    matchCount = matchCount + 1
                    Set match = candidateField
                End If
            End If
        End If
    Next candidateField
    If matchCount > 1 Then
        Err.Raise vbObjectError + 7568, "VisualTeX", _
            "The image Equation paragraph contains multiple visible number references."
    End If
    If matchCount = 1 Then Set VTImageEquationReferenceField = match
End Function

Public Function VTWriteImageEquationReferenceNumber( _
    ByVal documentObject As Document, _
    ByVal formulaRange As Range, _
    ByVal formulaId As String) As Range

    Dim formulaShape As InlineShape
    Dim formulaContainerRange As Range
    Dim paragraphRange As Range
    Dim suffixRange As Range
    Dim insertionRange As Range
    Dim fieldRange As Range
    Dim referenceField As Field
    Dim candidateField As Field
    Dim numberRange As Range
    Dim sequenceBookmarkName As String
    Dim numberBookmarkName As String
    Dim expectedNumber As String
    Dim formulaStart As Long
    Dim imageMetadata As String
    Dim imageFormulaReference As String
    Dim fieldStart As Long
    Dim fieldEnd As Long

    If documentObject Is Nothing Or formulaRange Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The image Equation REF target is invalid."
    End If
    If formulaRange.InlineShapes.Count <> 1 Or _
       formulaRange.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The image Equation REF formula is ambiguous."
    End If

    sequenceBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    numberBookmarkName = VTEquationNumberBookmarkName(formulaId)
    If Not documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The image Equation REF has no native sequence Bookmark."
    End If
    expectedNumber = VTEquationNumberTextForFormula( _
        documentObject, formulaId)
    If Len(expectedNumber) = 0 Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "The image Equation REF source has no visible number."
    End If

    Set formulaShape = formulaRange.InlineShapes(1)
    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range.Duplicate
    formulaStart = formulaShape.Range.Start
    imageMetadata = formulaShape.AlternativeText
    imageFormulaReference = formulaShape.Title
    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set suffixRange = documentObject.Range( _
        Start:=formulaContainerRange.End, End:=paragraphRange.End - 1)

    ' The formula paragraph is dedicated to VisualTeX. Only a previous managed
    ' REF/static/direct-SEQ number tail may be replaced here; unrelated fields or
    ' body text are never consumed.
    For Each candidateField In suffixRange.Fields
        If candidateField.Type = wdFieldRef Then
            If StrComp(VTReferenceTargetBookmarkName( _
               candidateField.Code.Text), sequenceBookmarkName, _
               vbTextCompare) <> 0 Then
                Err.Raise vbObjectError + 7560, "VisualTeX", _
                    "The image Equation tail contains an unrelated REF field."
            End If
        ElseIf Not VTIsNativeEquationSequenceField( _
           candidateField, VTNativeEquationLabelName()) Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The image Equation tail contains an unrelated Word field."
        End If
    Next candidateField
    If VTWordRangeHasMeaningfulText(suffixRange) Then
        Dim tailText As String
        tailText = suffixRange.Text
        For Each candidateField In suffixRange.Fields
            tailText = Replace$(tailText, candidateField.Result.Text, _
                "", 1, 1, vbBinaryCompare)
        Next candidateField
        tailText = Replace$(tailText, vbTab, "")
        tailText = Replace$(tailText, "(", "")
        tailText = Replace$(tailText, ")", "")
        tailText = Replace$(tailText, " ", "")
        tailText = Replace$(tailText, ChrW(160), "")
        tailText = Replace$(tailText, ChrW(8203), "")
        tailText = Replace$(tailText, ChrW(8288), "")
        If Len(tailText) <> 0 Then
            Err.Raise vbObjectError + 7560, "VisualTeX", _
                "The image Equation paragraph contains body text after the formula."
        End If
    End If

    If documentObject.Bookmarks.Exists(numberBookmarkName) Then
        documentObject.Bookmarks(numberBookmarkName).Delete
    End If
    If suffixRange.End > suffixRange.Start Then suffixRange.Delete

    Set formulaShape = VTResolveImageFormulaByIdentity( _
        documentObject, formulaId, imageMetadata, _
        imageFormulaReference, formulaStart)
    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set insertionRange = documentObject.Range( _
        Start:=formulaContainerRange.End, End:=formulaContainerRange.End)
    insertionRange.Text = vbTab & "()"

    Set formulaShape = VTResolveImageFormulaByIdentity( _
        documentObject, formulaId, imageMetadata, _
        imageFormulaReference, formulaStart)
    Set formulaContainerRange = _
        VTVisualTeXImageContainerRange(formulaShape)
    Set fieldRange = documentObject.Range( _
        Start:=formulaContainerRange.End + 2, _
        End:=formulaContainerRange.End + 2)
    Set referenceField = documentObject.Fields.Add( _
        Range:=fieldRange, Type:=wdFieldRef, _
        Text:=VTVisibleEquationReferenceFieldText(sequenceBookmarkName), _
        PreserveFormatting:=False)
    referenceField.Update
    If VTComparableEquationNumberText( _
       VTEquationSequenceResultText(referenceField)) <> _
       VTComparableEquationNumberText(expectedNumber) Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "Word did not resolve the visible image Equation REF" & _
            " [result=" & referenceField.Result.Text & _
            "; expected=" & expectedNumber & "]."
    End If

    fieldStart = VTEquationFieldStart(referenceField)
    fieldEnd = VTEquationFieldEnd(referenceField)
    Set numberRange = documentObject.Range( _
        Start:=fieldStart - 1, End:=fieldEnd + 1)
    If numberRange.Text <> "(" & expectedNumber & ")" Then
        Err.Raise vbObjectError + 7560, "VisualTeX", _
            "Word did not preserve the complete visible image Equation number."
    End If
    VTSetEquationNumberBookmarkExact _
        documentObject, formulaId, numberRange
    VTFormatVisibleEquationReference _
        documentObject, referenceField, numberRange
    Set VTWriteImageEquationReferenceNumber = numberRange.Duplicate
End Function

Public Function VTNativeEquationArrayReferenceField( _
    ByVal formulaRange As Range, _
    ByVal formulaId As String) As Field

    Dim exactRange As Range
    Dim candidateField As Field
    Dim targetBookmarkName As String
    Dim candidateTargetName As String
    Dim fieldStart As Long
    Dim fieldEnd As Long
    Dim matchCount As Long
    Dim match As Field

    If formulaRange Is Nothing Then Exit Function
    If formulaRange.OMaths.Count <> 1 Or _
       Not VTIsCanonicalUuid(formulaId) Then Exit Function
    Set exactRange = VTNativeMathRoot(formulaRange.OMaths(1)).Range.Duplicate
    targetBookmarkName = _
        VTEquationSequenceNumberBookmarkName(formulaId)
    ' The visible number REF must live inside this one OMath. Search the local
    ' formula Range instead of enumerating every Field in the document on each
    ' resolve; native numbering calls this resolver several times per operation.
    For Each candidateField In exactRange.Fields
        If candidateField.Type = wdFieldRef Then
            fieldStart = VTEquationFieldStart(candidateField)
            fieldEnd = VTEquationFieldEnd(candidateField)
            If fieldStart >= exactRange.Start And fieldEnd <= exactRange.End Then
                candidateTargetName = VTReferenceTargetBookmarkName( _
                    candidateField.Code.Text)
                If StrComp(candidateTargetName, targetBookmarkName, _
                   vbTextCompare) = 0 Then
                    matchCount = matchCount + 1
                    Set match = candidateField
                End If
            End If
        End If
    Next candidateField
    If matchCount > 1 Then
        Err.Raise vbObjectError + 7563, "VisualTeX", _
            "The native Equation array contains multiple visible number references."
    End If
    If matchCount = 1 Then Set VTNativeEquationArrayReferenceField = match
End Function

Public Sub VTFormatBodyEquationReference( _
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
        If .Hidden <> False Then .Hidden = False
        If .Color <> wdColorAutomatic Then .Color = wdColorAutomatic
        If .Position <> 0 Then .Position = 0
        If Abs(.Size - visibleSize) > 0.1 Then .Size = visibleSize
    End With
    With referenceRange.Font
        If .Hidden <> False Then .Hidden = False
        If .Color <> wdColorAutomatic Then .Color = wdColorAutomatic
        If .Position <> 0 Then .Position = 0
        If Abs(.Size - visibleSize) > 0.1 Then .Size = visibleSize
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

Public Function VTBodyEquationReferenceFormattingMatches( _
    ByVal documentObject As Document, _
    ByVal referenceRange As Range) As Boolean

    Dim visibleSize As Single
    If documentObject Is Nothing Or referenceRange Is Nothing Then Exit Function
    visibleSize = VTVisibleEquationNumberFontSize(documentObject)
    With referenceRange.Font
        VTBodyEquationReferenceFormattingMatches = _
            (.Hidden = False And .Color = wdColorAutomatic And _
             .Position = 0 And Abs(.Size - visibleSize) <= 0.1)
    End With
End Function

Private Sub VTNormalizeBodyEquationReferenceVisibilityInRange( _
    ByVal documentObject As Document, _
    ByVal scanRange As Range, _
    Optional ByVal changedFrom As Long = -1, _
    Optional ByVal updateFields As Boolean = False)

    Dim candidateField As Field
    Dim targetBookmarkName As String
    Dim sequenceBookmarkName As String
    Dim formulaId As String
    Dim targetKind As String
    Dim formulaIds As Variant
    Dim formulaIdsLoaded As Boolean
    Dim targetStart As Long
    Dim isVisualTeXReference As Boolean
    Dim shouldUpdateField As Boolean
    Dim shouldFormatVisualTeX As Boolean

    If documentObject Is Nothing Or scanRange Is Nothing Then Exit Sub
    For Each candidateField In scanRange.Fields
        If candidateField.Type = wdFieldRef And _
           Not candidateField.Result.Information(wdWithInTable) Then
            targetBookmarkName = VTReferenceTargetBookmarkName( _
                candidateField.Code.Text)
            formulaId = VTFormulaIdFromSequenceBookmarkName( _
                targetBookmarkName)
            isVisualTeXReference = (Len(formulaId) > 0)
            If Not isVisualTeXReference And _
               Left$(targetBookmarkName, 4) = "_Ref" Then
                If Not formulaIdsLoaded Then
                    formulaIds = VTValidNumberedFormulaIds(documentObject)
                    formulaIdsLoaded = True
                End If
                formulaId = VTFormulaIdForReferenceTarget( _
                    documentObject, targetBookmarkName, targetKind, _
                    candidateField.Result.Text, formulaIds)
                isVisualTeXReference = (Len(formulaId) > 0)
            End If

            targetStart = -1
            If isVisualTeXReference Then
                sequenceBookmarkName = _
                    VTEquationSequenceNumberBookmarkName(formulaId)
                If documentObject.Bookmarks.Exists(sequenceBookmarkName) Then
                    targetStart = documentObject.Bookmarks( _
                        sequenceBookmarkName).Range.Start
                End If
            ElseIf documentObject.Bookmarks.Exists(targetBookmarkName) Then
                ' Preserve Word-native Equation cross-references as well as
                ' VisualTeX references. Their private target Bookmark still
                ' provides a stable affected-position boundary.
                targetStart = documentObject.Bookmarks( _
                    targetBookmarkName).Range.Start
            End If

            shouldUpdateField = updateFields
            If shouldUpdateField And changedFrom >= 0 Then
                shouldUpdateField = _
                    (targetStart >= changedFrom)
            End If
            If shouldUpdateField Then candidateField.Update

            shouldFormatVisualTeX = isVisualTeXReference
            If shouldFormatVisualTeX And changedFrom >= 0 Then
                shouldFormatVisualTeX = _
                    (targetStart >= changedFrom)
            End If
            If shouldFormatVisualTeX Then
                VTFormatBodyEquationReference _
                    documentObject, candidateField, _
                    candidateField.Result.Duplicate
            End If
        End If
    Next candidateField
End Sub

Public Sub VTFormatVisibleEquationReference( _
    ByVal documentObject As Document, _
    ByVal referenceField As Field, _
    ByVal numberRange As Range, _
    Optional ByVal fontName As String = "", _
    Optional ByVal fontSize As Single = 0!)

    Dim visibleSize As Single
    Dim resultRange As Range

    If documentObject Is Nothing Or referenceField Is Nothing Or _
       numberRange Is Nothing Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The visible Equation number formatting target is missing."
    End If
    With numberRange.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBefore = 0!
        .SpaceAfter = 0!
    End With
    ' An image number in its MACROBUTTON shows the button's display text:
    ' Word reports the field's Result elsewhere (with the paragraph's own
    ' font), so the visible number is the range formatted and checked.
    If VTIsImagePlaceRef(referenceField) Then
        Set resultRange = numberRange
    Else
        Set resultRange = referenceField.Result
    End If
    If VT_WORD_KEEP_NUMBER_FORMAT Then
        ' Caret and open-time repairs keep the number's size, font and colour.
        resultRange.Font.Hidden = False
        resultRange.Font.Position = 0
        numberRange.Font.Hidden = False
        numberRange.Font.Position = 0
        If InStr(1, referenceField.Code.Text, "MERGEFORMAT", vbTextCompare) > 0 Or _
           resultRange.Font.Hidden <> False Then
            Err.Raise vbObjectError + 7549, "VisualTeX", _
                "The visible Equation number inherited helper formatting" & _
                " [code=" & referenceField.Code.Text & "]."
        End If
        Exit Sub
    End If
    If Len(fontName) = 0 Then fontName = VT_WORD_EQUATION_NUMBER_FONT_NAME
    visibleSize = fontSize
    If visibleSize <= 0! Or visibleSize > CSng(VT_WORD_MAX_FORMULA_FONT_SIZE_PT) Then _
        visibleSize = VTVisibleEquationNumberFontSize(documentObject)
    With resultRange.Font
        .Hidden = False
        .Color = wdColorAutomatic
        .Name = fontName
        .NameAscii = fontName
        .NameOther = fontName
        .Position = 0
        .Size = visibleSize
    End With
    With numberRange.Font
        .Hidden = False
        .Color = wdColorAutomatic
        .Name = fontName
        .NameAscii = fontName
        .NameOther = fontName
        .Position = 0
        .Size = visibleSize
    End With
    If InStr(1, referenceField.Code.Text, "MERGEFORMAT", _
       vbTextCompare) > 0 Or _
       resultRange.Font.Hidden <> False Or _
       resultRange.Font.Color <> wdColorAutomatic Or _
       StrComp(CStr(resultRange.Font.NameAscii), _
           fontName, vbTextCompare) <> 0 Or _
       StrComp(CStr(resultRange.Font.NameOther), _
           fontName, vbTextCompare) <> 0 Or _
       Abs(resultRange.Font.Size - visibleSize) > 0.1 Then
        Err.Raise vbObjectError + 7549, "VisualTeX", _
            "The visible Equation number inherited helper formatting" & _
            " [code=" & referenceField.Code.Text & _
            "; ascii=" & CStr(resultRange.Font.NameAscii) & _
            "; other=" & CStr(resultRange.Font.NameOther) & _
            "; size=" & CStr(resultRange.Font.Size) & _
            "; expectedSize=" & CStr(visibleSize) & "]."
    End If
End Sub

' Word for Mac walks Document.OMaths, and any Range.OMaths that spans many
' equations, in roughly cubic time: enumerating 406 equations takes about 8 s
' before a single property is read. A paragraph's own OMaths collection is
' local, so visiting the paragraphs of a Range and the equations inside each
' one returns the same equations, clipped to the same Range, in document order
' at a linear cost.
' A reference to a numbered Word equation (a REF to its number) shows that number as math; it is not a formula.
Public Function VTIsEquationReferenceMath(ByVal formulaMath As OMath) As Boolean
    Dim candidate As Field
    Dim mathRange As Range

    On Error GoTo NotReference
    Set mathRange = formulaMath.Range
    For Each candidate In mathRange.Paragraphs(1).Range.Fields
        If candidate.Type = wdFieldRef Then
            If candidate.Result.Start <= mathRange.Start And candidate.Result.End >= mathRange.End Then
                VTIsEquationReferenceMath = True
                Exit Function
            End If
        End If
    Next candidate
NotReference:
End Function
