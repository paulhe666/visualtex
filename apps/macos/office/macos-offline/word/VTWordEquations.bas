Attribute VB_Name = "VTWordEquations"
Option Explicit
Option Private Module

#Const VT_WORD_DIAGNOSTICS = False

' Word equations (OMML): inserting and joining them, their math roots and bookmarks, structure signatures,
' rollback documents, and display and inline layout.

Private Const VT_WORD_NATIVE_SIGNATURE_VARIABLE_PREFIX As String = _
    "VT_NativeSignature_"
Private VT_WORD_NATIVE_ROLLBACK_DOCUMENT As Document

Public Sub VTClosePreparedNativeRollbackDocument()
    On Error Resume Next
    If Not VT_WORD_NATIVE_ROLLBACK_DOCUMENT Is Nothing Then
        VT_WORD_NATIVE_ROLLBACK_DOCUMENT.Saved = True
        VT_WORD_NATIVE_ROLLBACK_DOCUMENT.Close _
            SaveChanges:=wdDoNotSaveChanges
    End If
    Set VT_WORD_NATIVE_ROLLBACK_DOCUMENT = Nothing
    On Error GoTo 0
End Sub

Public Sub VTPrepareNativeRollbackDocument( _
    ByVal sourceDocument As Document)

    On Error GoTo PreparationFailed
    VTClosePreparedNativeRollbackDocument
    Set VT_WORD_NATIVE_ROLLBACK_DOCUMENT = _
        Documents.Add(Visible:=False)
    VT_WORD_NATIVE_ROLLBACK_DOCUMENT.Saved = True
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    Exit Sub

PreparationFailed:
    ' Prewarming is only a latency optimization. Apply retains the original
    ' synchronous Documents.Add fallback when Word cannot prepare this hidden
    ' rollback container while the editor is opening.
    VTClosePreparedNativeRollbackDocument
    On Error Resume Next
    If Not sourceDocument Is Nothing Then sourceDocument.Activate
    On Error GoTo 0
End Sub

Public Function VTTakePreparedNativeRollbackDocument() As Document
    Dim preparedDocument As Document

    On Error GoTo PreparedDocumentInvalid
    If VT_WORD_NATIVE_ROLLBACK_DOCUMENT Is Nothing Then Exit Function
    Set preparedDocument = VT_WORD_NATIVE_ROLLBACK_DOCUMENT
    If Len(preparedDocument.Name) = 0 Then GoTo PreparedDocumentInvalid
    Set VT_WORD_NATIVE_ROLLBACK_DOCUMENT = Nothing
    Set VTTakePreparedNativeRollbackDocument = preparedDocument
    Exit Function

PreparedDocumentInvalid:
    VTClosePreparedNativeRollbackDocument
End Function

Public Function VTDocumentHasVisualTeXNativeBookmark( _
    ByVal documentObject As Document) As Boolean

    Dim candidate As Bookmark

    If documentObject Is Nothing Then Exit Function
    For Each candidate In documentObject.Bookmarks
        If Left$(candidate.Name, Len(VT_WORD_NATIVE_BOOKMARK_PREFIX)) = _
           VT_WORD_NATIVE_BOOKMARK_PREFIX Then
            VTDocumentHasVisualTeXNativeBookmark = True
            Exit Function
        End If
    Next candidate
End Function

Public Function VTNativeEquationFormulaContentRange( _
    ByVal formulaRange As Range) As Range

    Dim exactRange As Range
    Dim markerRange As Range

    If formulaRange Is Nothing Then Exit Function
    If formulaRange.OMaths.Count <> 1 Then Exit Function
    Set exactRange = VTNativeMathRoot(formulaRange.OMaths(1)).Range.Duplicate
    Set markerRange = VTNativeEquationArrayMarkerRange(exactRange)
    If Not markerRange Is Nothing Then exactRange.End = markerRange.Start
    If exactRange.End <= exactRange.Start Then Exit Function
    Set VTNativeEquationFormulaContentRange = exactRange
End Function

Public Function VTNativeMathXmlBody(ByVal sourceXml As String, _
    ByRef bodyStart As Long, ByRef bodyEnd As Long) As String

    Dim tagStart As Long
    Dim nextCharacter As String

    tagStart = InStr(1, sourceXml, "<m:oMath", vbBinaryCompare)
    Do While tagStart > 0
        nextCharacter = Mid$(sourceXml, tagStart + Len("<m:oMath"), 1)
        If nextCharacter = ">" Or nextCharacter = " " Then Exit Do
        tagStart = InStr(tagStart + 1, sourceXml, "<m:oMath", vbBinaryCompare)
    Loop
    If tagStart = 0 Then GoTo InvalidMath
    bodyStart = InStr(tagStart, sourceXml, ">", vbBinaryCompare) + 1
    bodyEnd = InStr(bodyStart, sourceXml, "</m:oMath>", vbBinaryCompare)
    If bodyStart < 2 Or bodyEnd <= bodyStart Then GoTo InvalidMath
    VTNativeMathXmlBody = Mid$(sourceXml, bodyStart, bodyEnd - bodyStart)
    Exit Function
InvalidMath:
    Err.Raise vbObjectError + 7563, "VisualTeX", _
        "Word did not expose a complete professional OMath XML body."
End Function

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

Public Sub VTPlaceCaretAfterInlineNativeEquation( _
    ByVal equationRange As Range)

    Dim documentObject As Document
    Dim exactEquationRange As Range
    Dim paragraphRange As Range
    Dim caretRange As Range

    If equationRange Is Nothing Then Exit Sub
    VTDeleteTrailingInlineNativeSeparator equationRange
    If equationRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7530, "VisualTeX", _
            "The inline equation caret target is missing."
    End If

    Set exactEquationRange = VTNativeMathRoot(equationRange.OMaths(1)).Range.Duplicate
    Set documentObject = exactEquationRange.Document
    Set paragraphRange = VTWordParagraphContainingFormula(exactEquationRange)
    If paragraphRange Is Nothing Then
        Err.Raise vbObjectError + 7530, "VisualTeX", _
            "Word could not resolve the inline equation paragraph."
    End If
    If documentObject.Windows.Count = 0 Then
        Err.Raise vbObjectError + 7530, "VisualTeX", _
            "The inline equation document has no active Word window."
    End If
    documentObject.Activate
    Set caretRange = documentObject.Range( _
        Start:=exactEquationRange.End, End:=exactEquationRange.End)
    caretRange.Font.Position = 0
    caretRange.Select

    ' Word for Mac keeps a collapsed Range at OMath.Range.End inside the math
    ' zone even when ordinary text already exists before the formula. Two pure
    ' Range transactions were verified by the real host and were both absorbed
    ' into OMath. MoveRight supplies the same caret affinity as Word's native
    ' equation editor, without adding a character to the document.
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
    ' Leave the caret with Word's native affinity. Any later Range insertion
    ' establishes and consumes its own temporary boundary in one transaction;
    ' a finished formula never leaves an invisible character in the document.
End Sub

Public Function VTNativeInsertionBoundary(ByVal requested As Range) As Range
    Dim doc As Document
    Dim probe As Range
    Dim math As OMath
    Dim leftMath As OMath
    Dim boundary As Range
    Dim position As Long
    Dim probeStart As Long
    Dim probeEnd As Long
    Set doc = requested.Document
    Set boundary = requested.Duplicate
    boundary.Collapse wdCollapseStart
    position = boundary.Start
    probeStart = position - 1
    If probeStart < doc.Content.Start Then probeStart = doc.Content.Start
    probeEnd = position + 1
    If probeEnd > doc.Content.End Then probeEnd = doc.Content.End
    Set probe = doc.Range(probeStart, probeEnd)
    For Each math In probe.OMaths
        Set leftMath = VTNativeMathRoot(math)
        If leftMath.Range.Start < position And leftMath.Range.End >= position Then
            Exit For
        End If
        Set leftMath = Nothing
    Next math
    If leftMath Is Nothing Then
        Set VTNativeInsertionBoundary = boundary
        Exit Function
    End If
    doc.Activate
    Set boundary = doc.Range(leftMath.Range.End, leftMath.Range.End)
    boundary.Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
    Selection.TypeText Text:=ChrW(8288)
    Set boundary = doc.Range(Selection.Start - 1, Selection.Start)
    If boundary.Text <> ChrW(8288) Or boundary.OMaths.Count <> 0 Then
        Err.Raise vbObjectError + 7530, "VisualTeX", _
            "Word did not expose an ordinary insertion boundary outside OMath."
    End If
    ' Replace this exact ordinary-text character rather than collapse back into
    ' OMath. The caller consumes it with FormattedText or Text immediately.
    Set VTNativeInsertionBoundary = boundary.Duplicate
End Function

Public Function VTOMathsInDocumentOrder(ByVal scanRange As Range) As Collection
    Dim result As Collection
    Dim documentObject As Document
    Dim scanParagraph As Paragraph
    Dim paragraphRange As Range
    Dim localRange As Range
    Dim localStart As Long
    Dim localEnd As Long
    Dim candidateMath As OMath

    Set result = New Collection
    Set VTOMathsInDocumentOrder = result
    If scanRange Is Nothing Then Exit Function
    If scanRange.OMaths.Count = 0 Then Exit Function
    If scanRange.End <= scanRange.Start Then
        For Each candidateMath In scanRange.OMaths
            result.Add candidateMath
        Next candidateMath
        Exit Function
    End If
    Set documentObject = scanRange.Document
    For Each scanParagraph In scanRange.Paragraphs
        Set paragraphRange = scanParagraph.Range
        localStart = paragraphRange.Start
        localEnd = paragraphRange.End
        If localStart < scanRange.Start Then localStart = scanRange.Start
        If localEnd > scanRange.End Then localEnd = scanRange.End
        If localEnd > localStart Then
            Set localRange = documentObject.Range( _
                Start:=localStart, End:=localEnd)
            If localRange.OMaths.Count > 0 Then
                For Each candidateMath In localRange.OMaths
                    result.Add candidateMath
                Next candidateMath
            End If
        End If
    Next scanParagraph
End Function

' The root Range of the native equation that contains a field, found through
' the field's own paragraph, or Nothing when the field is outside equations.
Public Function VTFieldNativeMathRoot(ByVal candidate As Field) As Range
    Dim paragraphRange As Range
    Dim candidateMath As OMath
    Dim rootRange As Range
    Dim fieldStart As Long

    If candidate Is Nothing Then Exit Function
    fieldStart = VTEquationFieldStart(candidate)
    Set paragraphRange = candidate.Code.Paragraphs(1).Range
    If paragraphRange.OMaths.Count = 0 Then Exit Function
    For Each candidateMath In paragraphRange.OMaths
        Set rootRange = VTNativeMathRoot(candidateMath).Range
        If rootRange.Start <= fieldStart And rootRange.End > fieldStart Then
            Set VTFieldNativeMathRoot = rootRange.Duplicate
            Exit Function
        End If
    Next candidateMath
End Function

Public Function VTNativeMathRoot(ByVal scopedMath As OMath) As OMath
    Dim current As OMath
    Dim parentMath As OMath
    Dim depth As Long
    If scopedMath Is Nothing Then Exit Function
    Set current = scopedMath
    ' Range.OMaths returns an equation clipped to that Range on Mac Word,
    ' even with NestingLevel=0. ParentOMath exposes the full native equation.
    ' Always ask for the parent once; nested arguments then climb to the root.
    For depth = 1 To 64
        Set parentMath = Nothing
        On Error Resume Next
        Set parentMath = current.ParentOMath
        Err.Clear
        On Error GoTo 0
        If parentMath Is Nothing Then Exit For
        If parentMath.Range.Start = current.Range.Start And _
           parentMath.Range.End = current.Range.End Then
            Set current = parentMath
            Exit For
        End If
        If parentMath.Range.Start > current.Range.Start Or _
           parentMath.Range.End < current.Range.End Then Exit For
        Set current = parentMath
    Next depth
    Set VTNativeMathRoot = current
End Function

Public Function VTNativeMathRootRange(ByVal scopedRange As Range) As Range
    Dim math As OMath
    If scopedRange Is Nothing Then Exit Function
    If scopedRange.OMaths.Count <> 1 Then Err.Raise vbObjectError + 7473, , _
        "The native equation range is ambiguous."
    Set math = VTNativeMathRoot(scopedRange.OMaths(1))
    Set VTNativeMathRootRange = math.Range.Duplicate
End Function

Public Function VTAdjacentNativeMath(ByVal requested As Range) As OMath
    Dim probe As Range
    Dim math As OMath
    Dim full As OMath
    Dim first As Long
    Dim last As Long
    first = requested.Start - 1
    If first < requested.Document.Content.Start Then first = requested.Document.Content.Start
    last = requested.Start + 1
    If last > requested.Document.Content.End Then last = requested.Document.Content.End
    Set probe = requested.Document.Range(first, last)
    For Each math In probe.OMaths
        Set full = VTNativeMathRoot(math)
        If full.Range.Start < requested.Start And _
           full.Range.End = requested.Start Then
            Set VTAdjacentNativeMath = full
            Exit Function
        End If
    Next math
    For Each math In probe.OMaths
        Set full = VTNativeMathRoot(math)
        If full.Range.Start = requested.Start Then
            Set VTAdjacentNativeMath = full
            Exit Function
        End If
    Next math
End Function

Private Function VTJoinNativeMathAtBoundary( _
    ByVal sourceDocument As Document, _
    ByVal requested As Range, ByVal existing As OMath, _
    ByVal incoming As Range, ByVal fontSizePt As Single) As Range

    Dim doc As Document
    Dim parsed As Document
    Dim backup As Document
    Dim backupRange As Range
    Dim replacement As Range
    Dim merged As Range
    Dim packageXml As String
    Dim oldBody As String
    Dim newBody As String
    Dim bodyStart As Long
    Dim bodyEnd As Long
    Dim unusedStart As Long
    Dim unusedEnd As Long
    Dim originalStart As Long
    Dim originalLength As Long
    Dim mergedLength As Long
    Dim originalCount As Long
    Dim path As String
    Dim replaced As Boolean
    Dim operationStage As String
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    Set doc = sourceDocument
#If VT_WORD_DIAGNOSTICS Then
    VTNativeParityTrace "join-enter", doc, requested
#End If
    incoming.Font.Size = fontSizePt
    newBody = VTNativeMathXmlBody(incoming.WordOpenXML, unusedStart, unusedEnd)
    doc.Activate
    Set existing = VTNativeMathRoot(existing)
    Set replacement = existing.Range.Duplicate
    originalStart = replacement.Start
    originalLength = replacement.End - replacement.Start
    originalCount = doc.OMaths.Count
    operationStage = "export-original"
#If VT_WORD_DIAGNOSTICS Then
    VTNativeParityTrace "join-original", doc, replacement
#End If
    packageXml = replacement.WordOpenXML
    oldBody = VTNativeMathXmlBody(packageXml, bodyStart, bodyEnd)
    If requested.Start <= originalStart Then
        newBody = newBody & oldBody
    Else
        newBody = oldBody & newBody
    End If
    ' Word's Insert Equation command joins adjacent expressions into one math
    ' zone. Stage that same native tree, preserving the old subtree and all its
    ' formatting verbatim. Never import an OMath container into an OMath Range.
    packageXml = Left$(packageXml, bodyStart - 1) & newBody & Mid$(packageXml, bodyEnd)
    path = VTApplicationSupportRoot() & "/native-join-" & VTNewUuidV4() & ".xml"
    VTWriteTextAtomic path, packageXml
    operationStage = "parse-joined"
    Set parsed = Documents.Open(FileName:=path, ConfirmConversions:=False, _
        ReadOnly:=True, AddToRecentFiles:=False, Visible:=False)
    If parsed.OMaths.Count <> 1 Then Err.Raise vbObjectError + 7472, , _
        "Word did not parse the joined native equation."
    Set merged = VTNativeMathRoot(parsed.OMaths(1)).Range.Duplicate
    operationStage = "inline-joined"
    merged.OMaths(1).Type = wdOMathInline
    mergedLength = merged.End - merged.Start
    operationStage = "backup-original"
    Set backup = Documents.Add(Visible:=False)
    Set backupRange = backup.Content.Duplicate
    backupRange.Collapse wdCollapseStart
    operationStage = "copy-backup"
    backupRange.FormattedText = replacement.FormattedText
    Set backupRange = VTNativeMathRoot(backup.OMaths(1)).Range.Duplicate
#If VT_WORD_DIAGNOSTICS Then
    VTNativeParityTrace "join-backup", backup, backupRange
#End If
    doc.Activate
    Set replacement = doc.Range(originalStart, originalStart + originalLength)
    replaced = True
    operationStage = "replace-zone"
    replacement.FormattedText = merged.FormattedText
    Set replacement = doc.Range(originalStart, originalStart + mergedLength)
    If replacement.OMaths.Count <> 1 Or doc.OMaths.Count <> originalCount Then _
        Err.Raise vbObjectError + 7472, , "Word did not preserve the joined math zone."
    operationStage = "verify-zone"
#If VT_WORD_DIAGNOSTICS Then
    VTNativeParityTrace "join-replaced", doc, replacement
    VTNativeParityTrace "join-close-parsed", parsed
#End If
    operationStage = "close-joined"
    parsed.Close SaveChanges:=wdDoNotSaveChanges
#If VT_WORD_DIAGNOSTICS Then
    VTNativeParityTrace "join-close-backup", backup
#End If
    operationStage = "close-backup"
    backup.Close SaveChanges:=wdDoNotSaveChanges
    doc.Activate
    On Error Resume Next
    Kill path
    On Error GoTo Failed
#If VT_WORD_DIAGNOSTICS Then
    VTNativeParityTrace "join-return", doc
#End If
    Set VTJoinNativeMathAtBoundary = VTNativeMathRootRange(doc.Range(originalStart, originalStart + mergedLength))
    Exit Function
Failed:
    failureNumber = Err.Number
    failureDescription = operationStage & ": " & Err.Description
    On Error Resume Next
    If replaced And Not backupRange Is Nothing Then _
        replacement.FormattedText = backupRange.FormattedText
    If Not parsed Is Nothing Then parsed.Close SaveChanges:=wdDoNotSaveChanges
    If Not backup Is Nothing Then backup.Close SaveChanges:=wdDoNotSaveChanges
    If Len(path) > 0 Then Kill path
    doc.Activate
    On Error GoTo 0
    Err.Raise failureNumber, "VisualTeX native math transaction", failureDescription
End Function

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

Public Function VTPreferredNativeDisplayFontSize( _
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

Public Function VTInsertNativeEquationFromSharedDocument( _
    ByVal targetRange As Range, _
    ByVal sharedDocument As Document, _
    ByVal sharedEquationIndex As Long, _
    ByVal displayMode As String, _
    ByVal displaySizing As Boolean, _
    Optional ByRef joinedNative As Boolean = False, _
    Optional ByVal nativeFontSizePt As Double = 0#) As Range

    Dim targetDocument As Document
    Dim stagingEquationRange As Range
    Dim insertionRange As Range
    Dim probeRange As Range
    Dim candidateRange As Range
    Dim nativeEquation As OMath
    Dim candidateMath As OMath
    Dim failedMath As OMath
    Dim insertionStart As Long
    Dim stagingEquationLength As Long
    Dim probeStart As Long
    Dim probeEnd As Long
    Dim candidateDistance As Long
    Dim bestDistance As Long
    Dim matchCount As Long
    Dim preferredSize As Single
    Dim resolvedEquationStart As Long
    Dim conversionErrorNumber As Long
    Dim conversionErrorDescription As String
    Dim adjacent As OMath

    If targetRange Is Nothing Or sharedDocument Is Nothing Then
        Err.Raise vbObjectError + 7450, "VisualTeX", _
            "The shared native-equation insertion target is missing."
    End If
    If sharedEquationIndex < 1 Or _
       sharedEquationIndex > sharedDocument.OMaths.Count Then
        Err.Raise vbObjectError + 7451, "VisualTeX", _
            "The shared native-equation staging index is invalid."
    End If
    If displayMode <> "inline" And displayMode <> "block" Then
        Err.Raise vbObjectError + 7451, "VisualTeX", _
            "The native-equation display mode is invalid."
    End If

    Set targetDocument = targetRange.Document
    If displaySizing Or displayMode = "block" Then
        preferredSize = VTPreferredNativeDisplayFontSize(targetRange)
    Else
        preferredSize = VTPreferredEquationFontSize(targetRange, False)
    End If
    insertionStart = targetRange.Start
    Set stagingEquationRange = _
        sharedDocument.OMaths(sharedEquationIndex).Range.Duplicate
    stagingEquationLength = _
        stagingEquationRange.End - stagingEquationRange.Start
    If stagingEquationLength < 1 Then
        Err.Raise vbObjectError + 7451, "VisualTeX", _
            "The shared native-equation staging item is empty."
    End If

    On Error GoTo RollbackConversion
    If nativeFontSizePt > 0# Then preferredSize = CSng(nativeFontSizePt)
    If displayMode = "inline" Then
        Set adjacent = VTAdjacentNativeMath(targetRange)
        If Not adjacent Is Nothing Then
            Set VTInsertNativeEquationFromSharedDocument = _
                VTJoinNativeMathAtBoundary(targetDocument, targetRange, adjacent, stagingEquationRange, preferredSize)
            joinedNative = True
            Exit Function
        End If
    End If
    Set insertionRange = targetRange.Duplicate
    insertionRange.Collapse wdCollapseStart
    insertionRange.FormattedText = stagingEquationRange.FormattedText
    targetDocument.Activate

    ' Resolve only the insertion neighborhood. The legacy single-formula path
    ' snapshots and scans every OMath in the document so it can distinguish an
    ' arbitrary replacement target. Batch import always inserts at an empty
    ' collapsed cursor, so a local probe is both exact and dramatically faster.
    probeStart = insertionStart - 1
    If probeStart < 0 Then probeStart = 0
    probeEnd = insertionStart + stagingEquationLength + 8
    If probeEnd > targetDocument.Content.End Then _
        probeEnd = targetDocument.Content.End
    If probeEnd <= probeStart Then probeEnd = probeStart + 1
    Set probeRange = targetDocument.Range( _
        Start:=probeStart, End:=probeEnd)
    bestDistance = 2147483647
    For Each candidateMath In probeRange.OMaths
        If VTOMathHasMeaningfulContent(candidateMath) Then
            candidateDistance = Abs(candidateMath.Range.Start - insertionStart)
            If candidateDistance < bestDistance Then
                bestDistance = candidateDistance
                matchCount = 1
                Set nativeEquation = candidateMath
            ElseIf candidateDistance = bestDistance Then
                matchCount = matchCount + 1
            End If
        End If
    Next candidateMath
    If matchCount <> 1 Or nativeEquation Is Nothing Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "Word did not transfer exactly one native equation from the shared staging document."
    End If

    Set candidateRange = nativeEquation.Range.Duplicate
    resolvedEquationStart = candidateRange.Start
    VTV2ApplyEquationTypography candidateRange, preferredSize
    If displayMode = "inline" Then
        Set candidateRange = VTFinalizeInlineNativeEquation(candidateRange)
    Else
        ' Equation format v2: Windows' display layout (WordFormulaHostLayout);
        ' no OMath justification, paragraph style or tab-stop change.
        Set candidateRange = VTV2FinishDisplayEquation(candidateRange, preferredSize)
        Set candidateRange = VTResolveNativeEquationRange( _
            targetDocument, resolvedEquationStart, 16)
    End If
    Set VTInsertNativeEquationFromSharedDocument = candidateRange.Duplicate
    Exit Function

RollbackConversion:
    conversionErrorNumber = Err.Number
    conversionErrorDescription = Err.Description
    On Error Resume Next
    If Not adjacent Is Nothing Then
        ' The complete native join owns and verifies its own rollback.
    ElseIf Not nativeEquation Is Nothing Then
        nativeEquation.Range.Delete
    Else
        Set failedMath = VTNativeMathNearStart( _
            targetDocument, insertionStart, stagingEquationLength + 8)
        If Not failedMath Is Nothing Then
            If Abs(failedMath.Range.Start - insertionStart) <= 2 Then _
                failedMath.Range.Delete
        End If
    End If
    On Error GoTo 0
    Err.Raise conversionErrorNumber, _
        "VisualTeX Word shared native equation insertion", _
        conversionErrorDescription
End Function

Public Function VTTryInsertNativeEquationFileFast( _
    ByVal targetRange As Range, _
    ByVal nativeDocumentPath As String, _
    ByVal displayMode As String, _
    ByVal displaySizing As Boolean, _
    Optional ByRef performancePhases As Variant, _
    Optional ByRef diagnosticStage As String) As Range

    Dim targetDocument As Document
    Dim insertionRange As Range
    Dim targetParagraph As Range
    Dim insertedRange As Range
    Dim probeRange As Range
    Dim formulaParagraph As Range
    Dim followingParagraph As Range
    Dim mergeRange As Range
    Dim equationRange As Range
    Dim candidateMath As OMath
    Dim nativeEquation As OMath
    Dim sourceShape As InlineShape
    Dim sourceBeforeRange As Range
    Dim sourceAfterRange As Range
    Dim insertionStart As Long
    Dim beforeContentEnd As Long
    Dim insertedLength As Long
    Dim probeEnd As Long
    Dim matchCount As Long
    Dim preferredSize As Single
    Dim phaseStartedAt As Single
    Dim insertSeconds As Double
    Dim resolveSeconds As Double
    Dim finalizeSeconds As Double
    Dim fastErrorNumber As Long
    Dim currentContentEnd As Long
    Dim sourceImageTarget As Boolean
    Dim paragraphHasMeaningfulText As Boolean
    Dim operationStage As String

    diagnosticStage = "validate-target"
    If targetRange Is Nothing Or Not displaySizing Then Exit Function
    If displayMode <> "inline" And displayMode <> "block" Then Exit Function
    If targetRange.Information(wdWithInTable) Then Exit Function
    If Len(nativeDocumentPath) = 0 Then Exit Function

    Set targetDocument = targetRange.Document
    Set insertionRange = targetRange.Duplicate
    insertionRange.Collapse wdCollapseStart
    If targetRange.InlineShapes.Count = 1 Then
        Set sourceShape = targetRange.InlineShapes(1)
        Set targetParagraph = VTWordParagraphContainingFormula( _
            sourceShape.Range.Duplicate)
        If targetParagraph Is Nothing Then
            diagnosticStage = "resolve-source-image-paragraph"
            Exit Function
        End If
        sourceImageTarget = True
    Else
        Set targetParagraph = insertionRange.Paragraphs(1).Range.Duplicate
        sourceImageTarget = False
    End If

    If sourceImageTarget Then
        Set sourceBeforeRange = targetDocument.Range( _
            Start:=targetParagraph.Start, End:=sourceShape.Range.Start)
        Set sourceAfterRange = targetDocument.Range( _
            Start:=sourceShape.Range.End, End:=targetParagraph.End)
        paragraphHasMeaningfulText = _
            VTWordRangeHasMeaningfulText(sourceBeforeRange) Or _
            VTWordRangeHasMeaningfulText(sourceAfterRange)
    Else
        paragraphHasMeaningfulText = _
            VTWordRangeHasMeaningfulText(targetParagraph)
    End If

    diagnosticStage = "validate-target-paragraph"
    If targetParagraph.Information(wdWithInTable) Or _
       targetParagraph.Fields.Count <> 0 Or _
       targetParagraph.OMaths.Count <> 0 Or _
       paragraphHasMeaningfulText Then
        diagnosticStage = "reject-target-paragraph"
        diagnosticStage = diagnosticStage & ";table=" & _
            CStr(targetParagraph.Information(wdWithInTable))
        diagnosticStage = diagnosticStage & ";fields=" & _
            CStr(targetParagraph.Fields.Count)
        diagnosticStage = diagnosticStage & ";maths=" & _
            CStr(targetParagraph.OMaths.Count)
        diagnosticStage = diagnosticStage & ";shapes=" & _
            CStr(targetParagraph.InlineShapes.Count)
        diagnosticStage = diagnosticStage & ";text=" & _
            CStr(paragraphHasMeaningfulText)
        Exit Function
    End If
    If sourceImageTarget Then
        diagnosticStage = "validate-source-image"
        If targetParagraph.InlineShapes.Count <> 1 Or _
           insertionRange.Start <> sourceShape.Range.Start Or _
           sourceShape.Range.Start <> _
               targetParagraph.InlineShapes(1).Range.Start Or _
           sourceShape.Range.End <> _
               targetParagraph.InlineShapes(1).Range.End Then Exit Function
    ElseIf targetParagraph.InlineShapes.Count <> 0 Then
        diagnosticStage = "reject-extra-inline-shapes"
        Exit Function
    End If

    If displaySizing Or displayMode = "block" Then
        preferredSize = VTPreferredNativeDisplayFontSize(targetRange)
    Else
        preferredSize = VTPreferredEquationFontSize(targetRange, False)
    End If
    insertionStart = insertionRange.Start
    beforeContentEnd = targetDocument.Content.End

    On Error GoTo FastInsertFailed
    operationStage = "insert-file"
    diagnosticStage = operationStage
    phaseStartedAt = Timer
    insertionRange.InsertFile _
        FileName:=nativeDocumentPath, _
        ConfirmConversions:=False, Link:=False, Attachment:=False
    insertSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    insertedLength = targetDocument.Content.End - beforeContentEnd
    If insertedLength <= 0 Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "Word did not insert native Equation file content."
    End If

    operationStage = "inspect-imported-range"
    diagnosticStage = operationStage
    phaseStartedAt = Timer
    Set insertedRange = targetDocument.Range( _
        Start:=insertionStart, End:=insertionStart + insertedLength)
    If insertedRange.Tables.Count <> 0 Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "The native Equation file inserted unsupported document objects."
    End If
    If sourceImageTarget Then
        If insertedRange.InlineShapes.Count > 1 Then
            Err.Raise vbObjectError + 7434, "VisualTeX", _
                "The native Equation file inserted an unexpected image beside the source formula."
        End If
        If insertedRange.InlineShapes.Count = 1 Then
            If sourceShape.Range.Start <> _
                   insertedRange.InlineShapes(1).Range.Start Or _
               sourceShape.Range.End <> _
                   insertedRange.InlineShapes(1).Range.End Then
                Err.Raise vbObjectError + 7434, "VisualTeX", _
                    "The native Equation file inserted an unrelated image."
            End If
        End If
    ElseIf insertedRange.InlineShapes.Count <> 0 Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "The native Equation file inserted an unexpected image."
    End If
    probeEnd = insertionStart + insertedLength + 1
    If probeEnd > targetDocument.Content.End Then _
        probeEnd = targetDocument.Content.End
    Set probeRange = targetDocument.Range( _
        Start:=insertionStart, End:=probeEnd)
    For Each candidateMath In probeRange.OMaths
        If VTOMathHasMeaningfulContent(candidateMath) Then
            matchCount = matchCount + 1
            Set nativeEquation = candidateMath
        End If
    Next candidateMath
    If matchCount <> 1 Or nativeEquation Is Nothing Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "The native Equation file did not insert exactly one local OMath."
    End If

    ' InsertFile imports the source DOCX paragraph mark as well as its OMath,
    ' leaving the equation in a newly imported paragraph immediately before the
    ' dedicated empty paragraph VisualTeX prepared for it. The imported paragraph
    ' renders correctly but Word can reject later InsertParagraphBefore calls on
    ' that Range. Merge it into the original empty target paragraph so the fast
    ' path preserves the same ordinary editable-paragraph semantics as the older
    ' FormattedText transfer path.
    operationStage = "inspect-imported-paragraph"
    diagnosticStage = operationStage
    Set formulaParagraph = nativeEquation.Range.Paragraphs(1).Range.Duplicate
    If formulaParagraph.End >= targetDocument.Content.End Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "The native Equation file did not preserve its target paragraph."
    End If
    Set followingParagraph = targetDocument.Range( _
        Start:=formulaParagraph.End, _
        End:=formulaParagraph.End).Paragraphs(1).Range.Duplicate
    If sourceImageTarget Then
        Set sourceBeforeRange = targetDocument.Range( _
            Start:=followingParagraph.Start, End:=sourceShape.Range.Start)
        Set sourceAfterRange = targetDocument.Range( _
            Start:=sourceShape.Range.End, End:=followingParagraph.End)
        paragraphHasMeaningfulText = _
            VTWordRangeHasMeaningfulText(sourceBeforeRange) Or _
            VTWordRangeHasMeaningfulText(sourceAfterRange)
    Else
        paragraphHasMeaningfulText = _
            VTWordRangeHasMeaningfulText(followingParagraph)
    End If
    If followingParagraph.Information(wdWithInTable) Or _
       followingParagraph.Fields.Count <> 0 Or _
       followingParagraph.OMaths.Count <> 0 Or _
       paragraphHasMeaningfulText Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "The native Equation file changed the prepared target paragraph."
    End If
    If sourceImageTarget Then
        If followingParagraph.InlineShapes.Count <> 1 Then
            Err.Raise vbObjectError + 7434, "VisualTeX", _
                "The native Equation file lost the pending source image."
        End If
        If sourceShape.Range.Start <> _
               followingParagraph.InlineShapes(1).Range.Start Or _
           sourceShape.Range.End <> _
               followingParagraph.InlineShapes(1).Range.End Then
            Err.Raise vbObjectError + 7434, "VisualTeX", _
                "The native Equation file changed the pending source image."
        End If
    ElseIf followingParagraph.InlineShapes.Count <> 0 Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "The native Equation file did not leave the prepared target paragraph empty."
    End If
    operationStage = "merge-imported-paragraph"
    diagnosticStage = operationStage
    Set mergeRange = targetDocument.Range( _
        Start:=formulaParagraph.End - 1, End:=formulaParagraph.End)
    If mergeRange.Text <> vbCr Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "The native Equation file paragraph boundary is invalid."
    End If
    mergeRange.Delete

    probeEnd = insertionStart + insertedLength + 1
    If probeEnd > targetDocument.Content.End Then _
        probeEnd = targetDocument.Content.End
    Set probeRange = targetDocument.Range( _
        Start:=insertionStart, End:=probeEnd)
    Set nativeEquation = Nothing
    matchCount = 0
    For Each candidateMath In probeRange.OMaths
        If VTOMathHasMeaningfulContent(candidateMath) Then
            matchCount = matchCount + 1
            Set nativeEquation = candidateMath
        End If
    Next candidateMath
    If matchCount <> 1 Or nativeEquation Is Nothing Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "Word lost the native Equation while normalizing its target paragraph."
    End If
    Set formulaParagraph = nativeEquation.Range.Paragraphs(1).Range.Duplicate
    If sourceImageTarget Then
        If formulaParagraph.InlineShapes.Count <> 1 Then
            Err.Raise vbObjectError + 7434, "VisualTeX", _
                "Word did not keep the pending source image beside the imported Equation."
        End If
        If sourceShape.Range.Start <> _
               formulaParagraph.InlineShapes(1).Range.Start Or _
           sourceShape.Range.End <> _
               formulaParagraph.InlineShapes(1).Range.End Then
            Err.Raise vbObjectError + 7434, "VisualTeX", _
                "Word moved the pending source image to another paragraph."
        End If
    ElseIf formulaParagraph.InlineShapes.Count <> 0 Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "Word left an unexpected image in the native Equation paragraph."
    End If
    resolveSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)

    operationStage = "finalize-equation"
    diagnosticStage = operationStage
    phaseStartedAt = Timer
    Set equationRange = nativeEquation.Range.Duplicate
    VTV2ApplyEquationTypography equationRange, preferredSize
    If displayMode = "inline" Then
        Set equationRange = _
            VTFinalizeInlineNativeEquation(equationRange)
    Else
        ' Equation format v2: Windows' display layout (WordFormulaHostLayout).
        Set equationRange = VTV2FinishDisplayEquation(equationRange, preferredSize)
    End If
    finalizeSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)

    If Not IsMissing(performancePhases) Then
        performancePhases = Array( _
            insertSeconds, resolveSeconds, finalizeSeconds)
    End If
    diagnosticStage = "PASS"
    Set VTTryInsertNativeEquationFileFast = equationRange.Duplicate
    Exit Function

FastInsertFailed:
    fastErrorNumber = Err.Number
    diagnosticStage = "failed:" & operationStage & ":" & Err.Description
    On Error Resume Next
    currentContentEnd = targetDocument.Content.End
    insertedLength = currentContentEnd - beforeContentEnd
    If insertedLength > 0 Then
        targetDocument.Range( _
            Start:=insertionStart, _
            End:=insertionStart + insertedLength).Delete
    End If
    If targetDocument.Content.End = beforeContentEnd Then
        Set targetParagraph = targetDocument.Range( _
            Start:=insertionStart, _
            End:=insertionStart).Paragraphs(1).Range.Duplicate
        If sourceImageTarget Then
            If targetParagraph.Fields.Count <> 0 Or _
               targetParagraph.OMaths.Count <> 0 Or _
               targetParagraph.InlineShapes.Count <> 1 Then
                On Error GoTo 0
                Err.Raise fastErrorNumber, _
                    "VisualTeX native Equation fast insertion", _
                    "Word could not restore the pending source-image paragraph."
            End If
            Set sourceShape = targetParagraph.InlineShapes(1)
            Set sourceBeforeRange = targetDocument.Range( _
                Start:=targetParagraph.Start, End:=sourceShape.Range.Start)
            Set sourceAfterRange = targetDocument.Range( _
                Start:=sourceShape.Range.End, End:=targetParagraph.End)
            If VTWordRangeHasMeaningfulText(sourceBeforeRange) Or _
               VTWordRangeHasMeaningfulText(sourceAfterRange) Then
                On Error GoTo 0
                Err.Raise fastErrorNumber, _
                    "VisualTeX native Equation fast insertion", _
                    "Word could not restore the pending source-image paragraph."
            End If
        Else
            VTNormalizePlainWordParagraph targetParagraph
        End If
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    Err.Raise fastErrorNumber, "VisualTeX native Equation fast insertion", _
        "Word could not roll back a failed direct native Equation insertion."
End Function

Public Function VTInsertNativeEquationAtRange( _
    ByVal targetRange As Range, _
    ByVal ommlBase64 As String, _
    ByVal nativeDocumentPath As String, _
    ByVal displayMode As String, _
    ByVal displaySizing As Boolean, _
    ByVal replaceTarget As Boolean, _
    Optional ByRef performancePhases As Variant, _
    Optional ByRef fastPathDiagnostic As String, _
    Optional ByVal replacementRollbackRange As Variant, _
    Optional ByRef joinedNative As Boolean = False, _
    Optional ByVal nativeFontSizePt As Double = 0#) As Range

    Dim ommlXml As String
    Dim insertionRange As Range
    Dim equationRange As Range
    Dim candidateRange As Range
    Dim probeRange As Range
    Dim targetDocument As Document
    Dim stagingDocument As Document
    Dim stagingEquationRange As Range
    Dim replacementBackupDocument As Document
    Dim replacementBackupRange As Range
    Dim replacementRange As Range
    Dim nativeEquation As OMath
    Dim originalReplacementMath As OMath
    Dim failedReplacementMath As OMath
    Dim candidateMath As OMath
    Dim existingMaths As Collection
    Dim insertionStart As Long
    Dim resolvedEquationStart As Long
    Dim preferredSize As Single
    Dim matchCount As Long
    Dim beforeMathCount As Long
    Dim stagingEquationLength As Long
    Dim candidateDistance As Long
    Dim bestDistance As Long
    Dim probeStart As Long
    Dim probeEnd As Long
    Dim mathIndex As Long
    Dim candidateEligible As Boolean
    Dim replacementApplied As Boolean
    Dim transferStarted As Boolean
    Dim conversionErrorNumber As Long
    Dim conversionErrorDescription As String
    Dim phaseStartedAt As Single
    Dim setupSeconds As Double
    Dim stagingSeconds As Double
    Dim transferSeconds As Double
    Dim resolveSeconds As Double
    Dim finalizeSeconds As Double
    Dim fastInsertPhases As Variant
    Dim fastInsertDiagnostic As String
    Dim adjacent As OMath
    Dim replacementWindow As Range

    fastPathDiagnostic = "not-attempted"
    phaseStartedAt = Timer
    If targetRange Is Nothing Then
        Err.Raise vbObjectError + 7450, "VisualTeX", "The native-equation insertion target is missing."
    End If
    Set targetDocument = targetRange.Document
#If VT_WORD_DIAGNOSTICS Then
    VTNativeParityTrace "insert-capture", targetDocument, targetRange
#End If
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

    ' Keep the entire fast-attempt -> rollback -> conservative fallback
    ' transaction under one error boundary. A failed direct InsertFile can make
    ' Word invalidate the caller's live Range even after the local helper restores
    ' document content; subsequent fallback setup must therefore be rolled back
    ' and reported with the fast-path diagnostic instead of escaping as raw 5904.
    On Error GoTo RollbackConversion

    ' A newly-created display equation always targets a dedicated empty
    ' paragraph. Word can import the minimal native DOCX directly into that
    ' paragraph far faster than opening the DOCX as a second document first.
    ' Keep replacements and inline-mixed content on the conservative staging
    ' path; the fast helper also rolls back and returns Nothing if its local
    ' structure checks do not hold on a particular Word build.
    If Not replaceTarget And displaySizing Then
        setupSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
        Set equationRange = VTTryInsertNativeEquationFileFast( _
            targetRange, nativeDocumentPath, displayMode, displaySizing, _
            fastInsertPhases, fastInsertDiagnostic)
        fastPathDiagnostic = fastInsertDiagnostic
        If Not equationRange Is Nothing Then
            If Not IsMissing(performancePhases) Then
                performancePhases = Array( _
                    setupSeconds, 0#, _
                    CDbl(fastInsertPhases(0)), _
                    CDbl(fastInsertPhases(1)), _
                    CDbl(fastInsertPhases(2)))
            End If
            fastPathDiagnostic = "PASS"
            Set VTInsertNativeEquationAtRange = equationRange.Duplicate
            Exit Function
        End If
        phaseStartedAt = Timer
    End If

    If displaySizing Or displayMode = "block" Then
        preferredSize = VTPreferredNativeDisplayFontSize(targetRange)
    Else
        preferredSize = VTPreferredEquationFontSize(targetRange, False)
    End If
    insertionStart = targetRange.Start
    Set existingMaths = New Collection
    ' Replacement must distinguish an arbitrary existing OMath from its new
    ' payload, so keep the full-document snapshot there. A normal insertion only
    ' mutates the insertion neighborhood; its pre-existing OMath snapshot is
    ' collected locally after the staging equation length is known.
    If replaceTarget Then
        ' The new OMath replaces the target in place, and the resolution probe
        ' below reaches at most one character before it and eight after it.
        ' Snapshot whole paragraphs around the target, so every OMath object is
        ' unclipped and keeps its identity, instead of Document.OMaths, which
        ' Word for Mac walks in cubic time (8 s at 406 equations).
        Set replacementWindow = VTParagraphWindow(targetRange, 1)
        For Each candidateMath In VTOMathsInDocumentOrder(replacementWindow)
            existingMaths.Add candidateMath
        Next candidateMath
        beforeMathCount = targetDocument.OMaths.Count
    End If
    Set insertionRange = targetRange.Duplicate
    If Not replaceTarget Then insertionRange.Collapse wdCollapseStart

    On Error GoTo RollbackConversion

    If replaceTarget Then
        If targetRange.OMaths.Count = 1 Then
            Set originalReplacementMath = targetRange.OMaths(1)
        End If
        ' VTCommitWordDispatch already owns an exact hidden-document backup for
        ' native edits so it can roll back failures that happen after insertion.
        ' Reuse that verified Range here instead of creating and copying the same
        ' OMath into a second hidden document on every Apply. Standalone callers
        ' retain the original self-contained backup path below.
        If IsObject(replacementRollbackRange) Then
            Set replacementBackupRange = replacementRollbackRange.Duplicate
        Else
            Set replacementBackupDocument = Documents.Add(Visible:=False)
            Set replacementBackupRange = replacementBackupDocument.Content
            replacementBackupRange.Collapse wdCollapseStart
            replacementBackupRange.FormattedText = targetRange.FormattedText
            If replacementBackupDocument.OMaths.Count = 1 Then
                Set replacementBackupRange = _
                    VTNativeMathRoot(replacementBackupDocument.OMaths(1)).Range.Duplicate
            ElseIf replacementBackupDocument.InlineShapes.Count = 1 Then
                Set replacementBackupRange = _
                    replacementBackupDocument.InlineShapes(1).Range.Duplicate
            Else
                Err.Raise vbObjectError + 7472, "VisualTeX", _
                    "Word could not back up the native-equation replacement target."
            End If
            targetDocument.Activate
        End If
    End If

    ' Word for Mac 16.89.1 rejects both a bare m:oMath node and a wrapped
' WordprocessingML document through Word's XML insertion API with runtime error
    ' 6145. The companion therefore creates a real minimal DOCX ZIP. Open
    ' that package and transfer Word's parsed OMath with FormattedText. This
    ' preserves structural fractions, n-ary operators, radicals, matrices and
    ' scripts without touching the clipboard or falling back to UnicodeMath.
    setupSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    phaseStartedAt = Timer
    Set stagingDocument = Documents.Open( _
        FileName:=nativeDocumentPath, _
        ConfirmConversions:=False, _
        ReadOnly:=True, _
        AddToRecentFiles:=False, _
        Visible:=False)
    If stagingDocument.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7434, "VisualTeX", "Word did not parse exactly one native equation from the OMML payload."
    End If
#If VT_WORD_DIAGNOSTICS Then
    VTNativeParityTrace "insert-stage-open", stagingDocument, targetRange
#End If
    Set stagingEquationRange = VTNativeMathRoot(stagingDocument.OMaths(1)).Range.Duplicate
    If nativeFontSizePt > 0# Then preferredSize = CSng(nativeFontSizePt)
    If Not replaceTarget And displayMode = "inline" Then
        targetDocument.Activate
        Set adjacent = VTAdjacentNativeMath(targetRange)
#If VT_WORD_DIAGNOSTICS Then
        VTNativeParityTrace "insert-adjacent=" & CStr(Not adjacent Is Nothing), targetDocument, targetRange
#End If
        stagingDocument.Activate
        If Not adjacent Is Nothing Then
            Set equationRange = VTJoinNativeMathAtBoundary( _
                targetDocument, targetRange, adjacent, stagingEquationRange, preferredSize)
#If VT_WORD_DIAGNOSTICS Then
            VTNativeParityTrace "insert-close-stage", stagingDocument, equationRange
#End If
            stagingDocument.Close SaveChanges:=wdDoNotSaveChanges
            Set stagingDocument = Nothing
            targetDocument.Activate
            joinedNative = True
            Set VTInsertNativeEquationAtRange = equationRange.Duplicate
            Exit Function
        End If
    End If
    stagingEquationLength = stagingEquationRange.End - stagingEquationRange.Start
    probeStart = insertionStart - 1
    If probeStart < 0 Then probeStart = 0
    probeEnd = insertionStart + stagingEquationLength + 8
    If probeEnd > targetDocument.Content.End Then _
        probeEnd = targetDocument.Content.End
    If Not replaceTarget Then
        Set probeRange = targetDocument.Range( _
            Start:=probeStart, End:=probeEnd)
        For Each candidateMath In probeRange.OMaths
            existingMaths.Add candidateMath
        Next candidateMath
    End If
    stagingSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    phaseStartedAt = Timer
    transferStarted = True
    insertionRange.FormattedText = stagingEquationRange.FormattedText
    replacementApplied = replaceTarget
    stagingDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set stagingDocument = Nothing
    targetDocument.Activate
    transferSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    phaseStartedAt = Timer

    ' FormattedText does not reliably expand the caller's Range on every Word
    ' for Mac build. Replacement keeps the conservative whole-document scan;
    ' ordinary insertion mirrors the validated shared-staging path and inspects
    ' only the neighborhood that could have changed.
    probeEnd = insertionStart + stagingEquationLength + 8
    If probeEnd > targetDocument.Content.End Then _
        probeEnd = targetDocument.Content.End
    If replaceTarget Then
        Set probeRange = replacementWindow.Duplicate
    Else
        Set probeRange = targetDocument.Range( _
            Start:=probeStart, End:=probeEnd)
    End If
    bestDistance = 2147483647
    For Each candidateMath In probeRange.OMaths
        candidateEligible = _
            Not VTOMathCollectionContains(existingMaths, candidateMath)
        If replaceTarget And Not originalReplacementMath Is Nothing Then
            If candidateMath Is originalReplacementMath Then
                candidateEligible = True
            End If
        End If
        If candidateEligible Then
            Set candidateRange = candidateMath.Range.Duplicate
            candidateDistance = Abs(candidateRange.Start - insertionStart)
            If candidateRange.End >= probeStart And _
               candidateRange.Start <= probeEnd And _
               VTOMathHasMeaningfulContent(candidateMath) Then
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

    For mathIndex = probeRange.OMaths.Count To 1 Step -1
        Set candidateMath = probeRange.OMaths(mathIndex)
        candidateEligible = _
            Not VTOMathCollectionContains(existingMaths, candidateMath)
        If replaceTarget And Not originalReplacementMath Is Nothing Then
            If candidateMath Is originalReplacementMath Then
                candidateEligible = True
            End If
        End If
        If candidateEligible Then
            Set candidateRange = candidateMath.Range.Duplicate
            If candidateRange.End >= probeStart And _
               candidateRange.Start <= probeEnd And _
               Not VTOMathHasMeaningfulContent(candidateMath) Then
                candidateRange.Delete
            End If
        End If
    Next mathIndex
    If replaceTarget Then
        If targetDocument.OMaths.Count <> beforeMathCount Then
            Err.Raise vbObjectError + 7434, "VisualTeX", _
                "Word did not replace the native equation as one clean OMath object."
        End If
    Else
        matchCount = 0
        For Each candidateMath In probeRange.OMaths
            candidateEligible = _
                Not VTOMathCollectionContains(existingMaths, candidateMath)
            If candidateEligible Then
                If VTOMathHasMeaningfulContent(candidateMath) Then
                    matchCount = matchCount + 1
                End If
            End If
        Next candidateMath
        If matchCount <> 1 Then
            Err.Raise vbObjectError + 7434, "VisualTeX", _
                "Word did not add exactly one clean local OMath object from the payload."
        End If
    End If
    resolveSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    phaseStartedAt = Timer

    Set equationRange = nativeEquation.Range.Duplicate
    resolvedEquationStart = equationRange.Start
    VTV2ApplyEquationTypography equationRange, preferredSize
    If displayMode = "inline" Then
        Set equationRange = _
            VTFinalizeInlineNativeEquation(equationRange)
    Else
        ' Equation format v2: Windows' display layout (WordFormulaHostLayout).
        Set equationRange = VTV2FinishDisplayEquation(equationRange, preferredSize)
        Set equationRange = VTResolveNativeEquationRange( _
            targetDocument, resolvedEquationStart, 16)
    End If

    finalizeSeconds = VTTimerElapsedSeconds(phaseStartedAt, Timer)
    If Not IsMissing(performancePhases) Then
        performancePhases = Array( _
            setupSeconds, stagingSeconds, transferSeconds, _
            resolveSeconds, finalizeSeconds)
    End If
    Set VTInsertNativeEquationAtRange = equationRange.Duplicate
    If Not replacementBackupDocument Is Nothing Then
        replacementBackupDocument.Close SaveChanges:=wdDoNotSaveChanges
        Set replacementBackupDocument = Nothing
    End If
    Exit Function

RollbackConversion:
    conversionErrorNumber = Err.Number
    conversionErrorDescription = Err.Description
    If Len(fastPathDiagnostic) > 0 And _
       fastPathDiagnostic <> "PASS" And _
       fastPathDiagnostic <> "not-attempted" Then
        conversionErrorDescription = _
            "fastPath=" & fastPathDiagnostic & "; fallback=" & _
            conversionErrorDescription
    End If
    On Error Resume Next
    If Not stagingDocument Is Nothing Then stagingDocument.Close SaveChanges:=wdDoNotSaveChanges
    If transferStarted Then
        Set probeRange = Nothing
        If replaceTarget Then
            If replacementWindow Is Nothing Then
                Set probeRange = targetDocument.Content.Duplicate
            Else
                Set probeRange = replacementWindow.Duplicate
            End If
        ElseIf stagingEquationLength > 0 Then
            probeStart = insertionStart - 1
            If probeStart < 0 Then probeStart = 0
            probeEnd = insertionStart + stagingEquationLength + 8
            If probeEnd > targetDocument.Content.End Then _
                probeEnd = targetDocument.Content.End
            Set probeRange = targetDocument.Range( _
                Start:=probeStart, End:=probeEnd)
        End If
        If Not probeRange Is Nothing Then
            For mathIndex = probeRange.OMaths.Count To 1 Step -1
                Set candidateMath = probeRange.OMaths(mathIndex)
                candidateEligible = _
                    Not VTOMathCollectionContains(existingMaths, candidateMath)
                If candidateEligible Then
                    Set candidateRange = candidateMath.Range.Duplicate
                    If candidateRange.End >= probeStart And _
                       candidateRange.Start <= probeEnd Then _
                        candidateRange.Delete
                End If
            Next mathIndex
        End If
    End If
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

Public Function VTPromoteNativeEquationToDisplay( _
    ByVal equationRange As Range) As Range

    Dim nativeEquation As OMath
    Dim exactRange As Range

    If equationRange Is Nothing Or equationRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7473, "VisualTeX", _
            "VisualTeX cannot promote a missing native equation to display mode."
    End If
    Set nativeEquation = equationRange.OMaths(1)
    ' Equation format v2: Windows' display layout at the equation's own size.
    Set exactRange = nativeEquation.Range.Duplicate
    VTWordPerformanceMark "native-promote-type-start"
    If exactRange.Font.Size > 0! And exactRange.Font.Size <= 1638! Then
        Set VTPromoteNativeEquationToDisplay = _
            VTV2FinishDisplayEquation(exactRange, exactRange.Font.Size)
    Else
        Set VTPromoteNativeEquationToDisplay = _
            VTV2FinishDisplayEquation(exactRange, 0#)
    End If
    VTWordPerformanceMark "native-promote-paragraph-complete"
End Function

Public Function VTFinalizeInlineNativeEquation( _
    ByVal equationRange As Range) As Range

    Dim documentObject As Document
    Dim nativeEquation As OMath
    Dim exactRange As Range
    Dim equationStart As Long

    If equationRange Is Nothing Or equationRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7475, "VisualTeX", _
            "VisualTeX cannot finalize a missing inline native equation."
    End If
    Set documentObject = equationRange.Document
    equationStart = equationRange.OMaths(1).Range.Start
    Set nativeEquation = equationRange.OMaths(1)
    nativeEquation.Type = wdOMathInline
    Set exactRange = VTResolveNativeEquationRange( _
        documentObject, equationStart, 16)
    exactRange.Font.Position = 0
    VTNormalizeInlineNativeParagraphAlignment exactRange
    VTDeleteTrailingInlineNativeSeparator exactRange
    Set exactRange = VTResolveNativeEquationRange( _
        documentObject, equationStart, 16)
    Set VTFinalizeInlineNativeEquation = exactRange.Duplicate
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

Public Function VTWordParagraphContainingFormula( _
    ByVal formulaRange As Range) As Range

    Dim exactRange As Range
    Dim probeRange As Range
    Dim paragraphRange As Range
    Dim paragraphEndProbe As Range
    Dim probeStart As Long

    If formulaRange Is Nothing Then Exit Function
    If formulaRange.OMaths.Count = 1 Then
        Set exactRange = VTNativeMathRoot(formulaRange.OMaths(1)).Range.Duplicate
    Else
        Set exactRange = formulaRange.Duplicate
    End If
    If exactRange.End <= exactRange.Start Then Exit Function

    ' A collapsed Range at OMath.Start can belong to the preceding paragraph on
    ' Word for Mac. Probe one character inside the formula so ordinary body text
    ' immediately above the insertion point can never be mistaken for the
    ' formula paragraph.
    probeStart = exactRange.Start + 1
    If probeStart >= exactRange.End Then probeStart = exactRange.Start
    Set probeRange = exactRange.Document.Range( _
        Start:=probeStart, End:=probeStart)
    Set paragraphRange = probeRange.Paragraphs(1).Range.Duplicate

    If exactRange.Start < paragraphRange.Start Or _
       exactRange.End > paragraphRange.End Then
        probeStart = exactRange.End - 1
        Set probeRange = exactRange.Document.Range( _
            Start:=probeStart, End:=probeStart)
        Set paragraphRange = probeRange.Paragraphs(1).Range.Duplicate
    End If

    ' Word for Mac can report Paragraphs.Count = 2 for a Range ending at the
    ' paragraph mark when the next paragraph contains a hidden SEQ field. Clamp
    ' the paragraph explicitly to the first real paragraph mark after its start
    ' so formula and helper paragraphs remain distinct regardless of that API
    ' quirk.
    Set paragraphEndProbe = exactRange.Document.Range( _
        Start:=paragraphRange.Start, End:=exactRange.Document.Content.End)
    With paragraphEndProbe.Find
        .ClearFormatting
        .Text = "^p"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
    End With
    If Not paragraphEndProbe.Find.Execute Then Exit Function
    paragraphRange.End = paragraphEndProbe.End

    If exactRange.Start < paragraphRange.Start Or _
       exactRange.End > paragraphRange.End Then Exit Function

    Set VTWordParagraphContainingFormula = paragraphRange
End Function

Public Function VTNativeMathNearStart( _
    ByVal documentObject As Document, _
    ByVal expectedStart As Long, _
    ByVal maximumDistance As Long) As OMath

    Dim candidateMath As OMath
    Dim match As OMath
    Dim candidateDistance As Long
    Dim bestDistance As Long
    Dim matchCount As Long
    Dim probe As Range
    Dim probeStart As Long
    Dim probeEnd As Long
    Dim contentStart As Long
    Dim contentEnd As Long

    If documentObject Is Nothing Or maximumDistance < 0 Then Exit Function
    ' Every candidate starts within maximumDistance of expectedStart, so a
    ' range spanning that window intersects all of them. Enumerating only its
    ' OMaths keeps the result identical while the cost no longer grows with
    ' the number of equations in the document (Document.OMaths made batch
    ' native insertion quadratic: 15 s per formula at 400 equations).
    contentStart = documentObject.Content.Start
    contentEnd = documentObject.Content.End
    probeStart = expectedStart - maximumDistance
    If probeStart < contentStart Then probeStart = contentStart
    probeEnd = expectedStart + maximumDistance + 1
    If probeEnd > contentEnd Then probeEnd = contentEnd
    If probeEnd <= probeStart Then Exit Function
    Set probe = documentObject.Range(Start:=probeStart, End:=probeEnd)
    bestDistance = 2147483647
    For Each candidateMath In probe.OMaths
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
    If matchCount = 1 Then Set VTNativeMathNearStart = VTNativeMathRoot(match)
End Function

Public Function VTResolveNativeEquationRange( _
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

' Whole paragraphs around centerRange: its own paragraphs plus extraParagraphs
' on each side. Equations never cross a paragraph mark, so every OMath of the
' window is complete.
Private Function VTParagraphWindow( _
    ByVal centerRange As Range, _
    ByVal extraParagraphs As Long) As Range

    Dim documentObject As Document
    Dim windowRange As Range
    Dim index As Long

    Set documentObject = centerRange.Document
    Set windowRange = centerRange.Duplicate
    windowRange.Start = windowRange.Paragraphs(1).Range.Start
    windowRange.End = windowRange.Paragraphs( _
        windowRange.Paragraphs.Count).Range.End
    For index = 1 To extraParagraphs
        If windowRange.Start > documentObject.Content.Start Then
            windowRange.Start = documentObject.Range( _
                Start:=windowRange.Start - 1, End:=windowRange.Start - 1) _
                .Paragraphs(1).Range.Start
        End If
        If windowRange.End < documentObject.Content.End Then
            windowRange.End = documentObject.Range( _
                Start:=windowRange.End, End:=windowRange.End) _
                .Paragraphs(1).Range.End
        End If
    Next index
    Set VTParagraphWindow = windowRange
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

Public Function VTOMathHasMeaningfulContent( _
    ByVal candidateMath As OMath) As Boolean

    Dim value As String

    If candidateMath Is Nothing Then Exit Function
    On Error GoTo EmptyMath
    value = candidateMath.Range.Text
    value = Replace$(value, vbCr, "")
    value = Replace$(value, vbLf, "")
    value = Replace$(value, vbTab, "")
    value = Replace$(value, Chr$(7), "")
    value = Replace$(value, ChrW(160), "")
    value = Replace$(value, ChrW(8203), "")
    value = Replace$(value, ChrW(8288), "")
    value = Replace$(value, ChrW(65279), "")
    VTOMathHasMeaningfulContent = (Len(Trim$(value)) > 0)
    Exit Function

EmptyMath:
    VTOMathHasMeaningfulContent = False
End Function

' True when candidateRange holds an equation with content. Word reports an
' empty OMath at the paragraph mark of an image formula that was converted from
' a display equation.
Public Function VTRangeHasMeaningfulMath(ByVal candidateRange As Range) As Boolean
    Dim candidateMath As OMath

    If candidateRange.OMaths.Count = 0 Then Exit Function
    For Each candidateMath In candidateRange.OMaths
        If VTOMathHasMeaningfulContent(candidateMath) Then
            VTRangeHasMeaningfulMath = True
            Exit Function
        End If
    Next candidateMath
End Function

Public Sub VTValidateOmmlFragment(ByVal ommlXml As String)
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

Public Function VTInsertCachedNativeEquationFast( _
    ByVal targetRange As Range, _
    ByVal nativeDocumentPath As String, _
    ByVal fontSizePt As Double) As Range

    Dim targetDocument As Document
    Dim stagingDocument As Document
    Dim stagingRange As Range
    Dim insertionRange As Range
    Dim probeRange As Range
    Dim equationRange As Range
    Dim nativeMath As OMath
    Dim insertionStart As Long
    Dim probeEnd As Long
    Dim stagingLength As Long
    Dim targetMutationStarted As Boolean
    Dim insertErrorNumber As Long
    Dim insertErrorDescription As String

    If targetRange Is Nothing Or Not VTValidWordFormulaFontSize(fontSizePt) Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "The cached native equation insertion target is invalid."
    End If
    VTValidateAbsoluteVisualTeXPath nativeDocumentPath
    ' The caller already performed the one sandbox-safe existence preflight.
    ' Avoid a second AppleScriptTask round trip here; Documents.Open remains the
    ' authoritative guard if the cache disappears between those two steps.
    Set targetDocument = targetRange.Document
    insertionStart = targetRange.Start
    On Error GoTo InsertFailed

    Set stagingDocument = Documents.Open( _
        FileName:=nativeDocumentPath, _
        ConfirmConversions:=False, _
        ReadOnly:=True, _
        AddToRecentFiles:=False, _
        Visible:=False)
    If stagingDocument.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "The cached native Word document does not contain exactly one equation."
    End If
    Set stagingRange = VTNativeMathRoot(stagingDocument.OMaths(1)).Range.Duplicate
    stagingLength = stagingRange.End - stagingRange.Start
    Set insertionRange = targetDocument.Range( _
        Start:=insertionStart, End:=insertionStart)
    targetMutationStarted = True
    insertionRange.FormattedText = stagingRange.FormattedText
    stagingDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set stagingDocument = Nothing
    targetDocument.Activate

    probeEnd = insertionStart + stagingLength + 8
    If probeEnd > targetDocument.Content.End Then
        probeEnd = targetDocument.Content.End
    End If
    Set probeRange = targetDocument.Range( _
        Start:=insertionStart, End:=probeEnd)
    If probeRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "Word did not create one cached native equation at the insertion point."
    End If
    Set nativeMath = VTNativeMathRoot(probeRange.OMaths(1))
    If Not VTOMathHasMeaningfulContent(nativeMath) Then
        Err.Raise vbObjectError + 7434, "VisualTeX", _
            "The cached native equation is empty."
    End If
    Set equationRange = nativeMath.Range.Duplicate
    VTV2ApplyEquationTypography equationRange, fontSizePt
    Set VTInsertCachedNativeEquationFast = equationRange
    Exit Function

InsertFailed:
    insertErrorNumber = Err.Number
    insertErrorDescription = Err.Description
    On Error Resume Next
    If Not stagingDocument Is Nothing Then
        stagingDocument.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If Not targetMutationStarted Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    If insertionStart >= 0 And Not targetDocument Is Nothing Then
        probeEnd = insertionStart + 64
        If probeEnd > targetDocument.Content.End Then
            probeEnd = targetDocument.Content.End
        End If
        Set probeRange = targetDocument.Range( _
            Start:=insertionStart, End:=probeEnd)
        If probeRange.OMaths.Count = 1 Then
            probeRange.OMaths(1).Range.Delete
        End If
    End If
    On Error GoTo 0
    Err.Raise insertErrorNumber, "VisualTeX cached native insertion", _
        insertErrorDescription
End Function

Public Function VTNativeWordDocumentPath(ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7434, "VisualTeX", "The native Word formula id is invalid."
    End If
    VTNativeWordDocumentPath = _
        VTApplicationSupportRoot() & "/NativeDocuments/" & formulaId & ".docx"
End Function

Public Function VTNativeFormulaBookmarkName(ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7457, "VisualTeX", "VisualTeX cannot bookmark a native equation with an invalid formula id."
    End If
    VTNativeFormulaBookmarkName = _
        VT_WORD_NATIVE_BOOKMARK_PREFIX & Replace$(formulaId, "-", "")
    If Len(VTNativeFormulaBookmarkName) > 40 Then
        Err.Raise vbObjectError + 7458, "VisualTeX", "VisualTeX generated a native equation bookmark longer than Word permits."
    End If
End Function

Public Function VTTryFormulaIdFromNativeBookmark( _
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

Public Sub VTSetNativeFormulaBookmark( _
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
    Set exactRange = VTNativeMathRoot(equationRange.OMaths(1)).Range.Duplicate
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

Public Function VTNativeMathForBookmark(ByVal nativeBookmark As Bookmark) As OMath
    Dim nearbyRange As Range
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
        Set VTNativeMathForBookmark = VTNativeMathRoot(nativeBookmark.Range.OMaths(1))
        Exit Function
    End If

    bookmarkStart = nativeBookmark.Range.Start
    bookmarkEnd = nativeBookmark.Range.End
    bestDistance = 1E+30
    ' Every candidate overlaps [bookmarkStart - 1, bookmarkEnd + 1], so the
    ' paragraphs around the Bookmark hold all of them.
    Set nearbyRange = nativeBookmark.Range.Document.Range( _
        Start:=bookmarkStart, End:=bookmarkEnd)
    If nearbyRange.Start > 0 Then nearbyRange.Start = nearbyRange.Start - 1
    If nearbyRange.End < nearbyRange.Document.Content.End Then _
        nearbyRange.End = nearbyRange.End + 1
    nearbyRange.Start = nearbyRange.Paragraphs(1).Range.Start
    nearbyRange.End = nearbyRange.Paragraphs(nearbyRange.Paragraphs.Count).Range.End
    For Each candidate In VTOMathsInDocumentOrder(nearbyRange)
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
    If matchCount = 1 Then Set VTNativeMathForBookmark = VTNativeMathRoot(match)
    Exit Function

NoMatch:
    Set VTNativeMathForBookmark = Nothing
End Function

Public Function VTTryResolveSingleNativeMath( _
    ByVal selectedRange As Range, _
    ByRef nativeMath As OMath) As Boolean

    Dim probeRange As Range

    Set nativeMath = Nothing
    If selectedRange Is Nothing Then Exit Function
    On Error GoTo ResolveFailed
    Set probeRange = selectedRange.Duplicate
    If probeRange.OMaths.Count <> 1 Then
        If probeRange.Start > 0 Then
            probeRange.MoveStart Unit:=wdCharacter, Count:=-1
        End If
        If probeRange.End < probeRange.Document.Content.End Then
            probeRange.MoveEnd Unit:=wdCharacter, Count:=1
        End If
    End If
    If probeRange.OMaths.Count <> 1 Then Exit Function
    Set nativeMath = VTNativeMathRoot(probeRange.OMaths(1))
    VTTryResolveSingleNativeMath = True
    Exit Function

ResolveFailed:
    Set nativeMath = Nothing
    Err.Clear
End Function

Public Function VTTryFindNativeFormulaBookmarkLocally( _
    ByVal selectedRange As Range, _
    ByRef resolvedBookmark As Bookmark) As Boolean

    Dim probeRange As Range
    Dim mathRange As Range
    Dim candidate As Bookmark
    Dim formulaId As String
    Dim match As Bookmark
    Dim matchCount As Long

    Set resolvedBookmark = Nothing
    If selectedRange Is Nothing Then Exit Function
    On Error GoTo LocalMiss

    Set probeRange = selectedRange.Duplicate
    If probeRange.OMaths.Count <> 1 Then
        probeRange.MoveStart Unit:=wdCharacter, Count:=-1
        probeRange.MoveEnd Unit:=wdCharacter, Count:=1
    End If
    If probeRange.OMaths.Count <> 1 Then Exit Function
    Set mathRange = VTNativeMathRoot(probeRange.OMaths(1)).Range.Duplicate

    ' Current VisualTeX OMML formulas persist one exact VT_F_ Bookmark around
    ' the equation. Range.Bookmarks limits lookup to the clicked equation and
    ' avoids walking every Bookmark in a formula-dense document.
    For Each candidate In mathRange.Bookmarks
        If VTTryFormulaIdFromNativeBookmark(candidate.Name, formulaId) Then
            If candidate.Range.Start <= mathRange.End And _
               candidate.Range.End >= mathRange.Start Then
                matchCount = matchCount + 1
                Set match = candidate
            End If
        End If
    Next candidate
    If matchCount = 1 Then
        Set resolvedBookmark = match
        VTTryFindNativeFormulaBookmarkLocally = True
    End If
    Exit Function

LocalMiss:
    Set resolvedBookmark = Nothing
    Err.Clear
End Function

Public Function VTFindNativeFormulaBookmark( _
    ByVal selectedRange As Range, _
    Optional ByVal requireMatch As Boolean = True) As Bookmark

    Dim candidate As Bookmark
    Dim match As Bookmark
    Dim formulaId As String
    Dim matchCount As Long
    Dim candidateMath As OMath

    If selectedRange Is Nothing Then GoTo NoMatch
    If VTTryFindNativeFormulaBookmarkLocally( _
       selectedRange, match) Then
        Set VTFindNativeFormulaBookmark = match
        Exit Function
    End If

    ' Legacy or damaged documents can lack the exact local Bookmark. Retain the
    ' full-document resolver only as a compatibility fallback; genuine current
    ' VisualTeX formulas never pay this formula-count-dependent cost.
    Set match = Nothing
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

Public Function VTWordNativeSignatureVariableName( _
    ByVal formulaId As String) As String

    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7467, "VisualTeX", _
            "VisualTeX cannot address a native formula signature for an invalid id."
    End If
    VTWordNativeSignatureVariableName = _
        VT_WORD_NATIVE_SIGNATURE_VARIABLE_PREFIX & _
        Replace$(formulaId, "-", "_")
End Function

Private Function VTNativeMathStructureSignature( _
    ByVal nativeMath As OMath) As String

    Dim sourceXml As String
    Dim mathStart As Long
    Dim mathEnd As Long
    Dim fragment As String
    Dim hashValue As Double
    Dim characterValue As Long
    Dim characterIndex As Long

    If nativeMath Is Nothing Then Exit Function
    sourceXml = nativeMath.Range.WordOpenXML
    mathStart = InStr(1, sourceXml, "<m:oMath", vbBinaryCompare)
    If mathStart = 0 Then Exit Function
    mathEnd = InStr(mathStart, sourceXml, "</m:oMath>", vbBinaryCompare)
    If mathEnd = 0 Then Exit Function
    fragment = Mid$( _
        sourceXml, mathStart, _
        mathEnd - mathStart + Len("</m:oMath>"))
    fragment = Replace$(fragment, vbCr, "")
    fragment = Replace$(fragment, vbLf, "")
    fragment = Replace$(fragment, vbTab, "")
    Do While InStr(1, fragment, "> ", vbBinaryCompare) > 0
        fragment = Replace$(fragment, "> ", ">")
    Loop
    Do While InStr(1, fragment, " <", vbBinaryCompare) > 0
        fragment = Replace$(fragment, " <", "<")
    Loop
    If Len(fragment) = 0 Then Exit Function

    hashValue = 7#
    For characterIndex = 1 To Len(fragment)
        characterValue = AscW(Mid$(fragment, characterIndex, 1))
        If characterValue < 0 Then characterValue = characterValue + 65536
        hashValue = hashValue * 131# + CDbl(characterValue)
        hashValue = hashValue - _
            Fix(hashValue / 2147483629#) * 2147483629#
    Next characterIndex
    VTNativeMathStructureSignature = _
        CStr(Len(fragment)) & "|" & Hex$(CLng(hashValue))
End Function

Public Function VTStableTextHash(ByVal value As String) As String
    Dim hashValue As Double
    Dim characterValue As Long
    Dim characterIndex As Long

    If Len(value) = 0 Then Exit Function
    hashValue = 7#
    For characterIndex = 1 To Len(value)
        characterValue = AscW(Mid$(value, characterIndex, 1))
        If characterValue < 0 Then characterValue = characterValue + 65536
        hashValue = hashValue * 131# + CDbl(characterValue)
        hashValue = hashValue - _
            Fix(hashValue / 2147483629#) * 2147483629#
    Next characterIndex
    VTStableTextHash = CStr(Len(value)) & "|" & Hex$(CLng(hashValue))
End Function

Public Function VTNativeMathFastSignature( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal nativeMath As OMath) As String

    Dim equationRange As Range
    Dim signatureRange As Range
    Dim formulaContentRange As Range
    Dim functionIndex As Long
    Dim functionTypes As String
    Dim signatureText As String
    Dim numberedText As String
    Dim numbered As Boolean

    If documentObject Is Nothing Or nativeMath Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then Exit Function

    Set equationRange = nativeMath.Range.Duplicate
    Set signatureRange = equationRange.Duplicate
    numbered = documentObject.Bookmarks.Exists( _
        VTEquationNumberBookmarkName(formulaId))
    If numbered Then
        numberedText = "1"
        ' A numbered native Equation contains its generated REF result in the
        ' outer OMath array. Renumbering must not make an unchanged mathematical
        ' body look edited, so hash only the formula side of the exact array.
        On Error Resume Next
        Set formulaContentRange = _
            VTNativeEquationFormulaContentRange(equationRange)
        Err.Clear
        On Error GoTo 0
        If Not formulaContentRange Is Nothing Then
            Set signatureRange = formulaContentRange.Duplicate
        End If
    Else
        numberedText = "0"
    End If

    signatureText = signatureRange.Text
    signatureText = Replace$(signatureText, vbCrLf, vbLf)
    signatureText = Replace$(signatureText, vbCr, vbLf)
    signatureText = Replace$(signatureText, Chr$(7), "")

    ' Range.Text is inexpensive on Word for Mac. Pair it with the ordered OMath
    ' function kinds so two structures with the same visible characters (for
    ' example a built-up fraction and a literal slash expression) cannot share
    ' a cache identity. This avoids serializing WordOpenXML on every Apply.
    On Error Resume Next
    For functionIndex = 1 To nativeMath.Functions.Count
        functionTypes = functionTypes & "," & _
            CStr(nativeMath.Functions(functionIndex).Type)
    Next functionIndex
    Err.Clear
    On Error GoTo 0

    VTNativeMathFastSignature = "v2:" & VTStableTextHash( _
        CStr(nativeMath.Type) & "|" & numberedText & "|" & _
        CStr(signatureRange.OMaths.Count) & "|" & _
        functionTypes & "|" & signatureText)
End Function

Public Sub VTSetWordNativeSignature( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal nativeMath As OMath)

    Dim signatureValue As String

    signatureValue = VTNativeMathFastSignature( _
        documentObject, formulaId, nativeMath)
    If Len(signatureValue) = 0 Then
        Err.Raise vbObjectError + 7467, "VisualTeX", _
            "Word did not expose a stable native formula structure."
    End If
    VTSetDocumentVariable _
        documentObject, _
        VTWordNativeSignatureVariableName(formulaId), _
        signatureValue
End Sub

Public Function VTWordNativeSignatureMatches( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal nativeMath As OMath) As Boolean

    Dim storedSignature As String
    Dim currentSignature As String

    If documentObject Is Nothing Or nativeMath Is Nothing Then Exit Function
    If Not VTTryGetDocumentVariable( _
       documentObject, VTWordNativeSignatureVariableName(formulaId), _
       storedSignature) Then Exit Function
    currentSignature = VTNativeMathFastSignature( _
        documentObject, formulaId, nativeMath)
    If Len(currentSignature) = 0 Then Exit Function

    If Left$(storedSignature, 3) = "v2:" Then
        VTWordNativeSignatureMatches = _
            (StrComp(storedSignature, currentSignature, _
                vbBinaryCompare) = 0)
        Exit Function
    End If

    ' Existing 1.2.5 documents retain the former OpenXML-derived hash. Pay that
    ' serialization cost once when such a document is first inspected, then
    ' migrate its variable to the inexpensive v2 signature.
    If StrComp(storedSignature, _
       VTNativeMathStructureSignature(nativeMath), _
       vbBinaryCompare) = 0 Then
        VTWordNativeSignatureMatches = True
        On Error Resume Next
        VTSetDocumentVariable documentObject, _
            VTWordNativeSignatureVariableName(formulaId), currentSignature
        Err.Clear
        On Error GoTo 0
    End If
End Function

' createdNow: formulaId was made by this very commit, so no reference to its bookmarks can exist yet and the scan of every field
' of the document for one (60 ms in a document of 400 equations) is left out. The ids of equations it replaced are scanned as before.
Public Sub VTFinishPureNativeEquation( _
    ByVal exact As Range, ByVal formulaId As String, _
    Optional ByVal createdNow As Boolean = False)

    Dim ids As New Collection
    Dim mark As Bookmark
    Dim id As Variant
    Dim oldId As String
    Dim doc As Document
    Dim tracked As Range
    Dim first As Long
    Dim last As Long
    If exact Is Nothing Then Exit Sub
    Set doc = exact.Document
    Set tracked = exact.Duplicate
    first = exact.Start
    last = exact.End
    ids.Add formulaId
    For Each mark In doc.Bookmarks
        If Left$(mark.Name, Len(VT_WORD_NATIVE_BOOKMARK_PREFIX)) = VT_WORD_NATIVE_BOOKMARK_PREFIX Then
            If mark.Range.Start >= first And mark.Range.End <= last Then
                oldId = VTFormulaIdFromBookmarkSuffix(Mid$(mark.Name, Len(VT_WORD_NATIVE_BOOKMARK_PREFIX) + 1))
                If oldId <> formulaId And VTIsCanonicalUuid(oldId) Then ids.Add oldId
            End If
        End If
    Next mark
    For Each id In ids
        Set exact = VTNativeMathRoot(tracked.OMaths(1)).Range.Duplicate
        VTRetireLegacyNativeBoundary exact, CStr(id)
        VTFinishPureNativeIdentity exact, CStr(id), _
            (createdNow And StrComp(CStr(id), formulaId, vbBinaryCompare) = 0)
    Next id
End Sub

Private Sub VTFinishPureNativeIdentity( _
    ByVal exact As Range, ByVal formulaId As String, _
    Optional ByVal noReferencesYet As Boolean = False)

    Dim doc As Document
    Dim prefixes As Variant
    Dim prefix As Variant
    Dim oldName As String
    Dim newName As String
    Dim source As Range
    Dim candidate As Field
    Dim numberField As Field
    Dim referenced As Boolean
    If exact Is Nothing Then Exit Sub
    If exact.OMaths.Count <> 1 Then Err.Raise vbObjectError + 7606
    Set doc = exact.Document
    Set numberField = VTPureNativeNumberField(exact)
        prefixes = Array(VT_WORD_NUMBER_BOOKMARK_PREFIX, _
            VT_WORD_SEQUENCE_NUMBER_BOOKMARK_PREFIX, _
            VT_WORD_CAPTION_BOOKMARK_PREFIX, _
            VT_WORD_NATIVE_BOOKMARK_PREFIX, _
            VT_WORD_IMAGE_OWNER_BOOKMARK_PREFIX)
    For Each prefix In prefixes
        oldName = CStr(prefix) & Replace$(formulaId, "-", "")
        If doc.Bookmarks.Exists(oldName) Then
            referenced = False
            If Not noReferencesYet Then
                For Each candidate In doc.Fields
                    If candidate.Type = wdFieldRef Then
                        If StrComp(VTReferenceTargetBookmarkName(candidate.Code.Text), _
                            oldName, vbTextCompare) = 0 Then referenced = True
                    End If
                Next candidate
            End If
            If referenced Then
                If numberField Is Nothing Then
                    Set source = doc.Bookmarks(oldName).Range.Duplicate
                Else
                    Set source = VTPureNativeVisibleNumberRange( _
                        exact, numberField)
                End If
                ' A Word-native cross-reference bookmark is created only when an
                ' existing body reference needs it. Never leave a VT identity.
                newName = "_Ref" & Replace$(VTNewUuidV4(), "-", "")
                doc.Bookmarks.Add Name:=newName, Range:=source
                For Each candidate In doc.Fields
                    If candidate.Type = wdFieldRef Then
                        If StrComp(VTReferenceTargetBookmarkName(candidate.Code.Text), _
                            oldName, vbTextCompare) = 0 Then
                            candidate.Code.Text = Replace$(candidate.Code.Text, oldName, newName)
                            candidate.Update
                        End If
                    End If
                Next candidate
            End If
            doc.Bookmarks(oldName).Delete
        End If
    Next prefix
        VTDeleteWordLatexPayload doc, formulaId
        VTDeleteWordOmmlPayload doc, formulaId
        VTDeleteWordMetadataPayload doc, formulaId
        VTDeleteDocumentVariable doc, VTWordFormatVariableName(formulaId)
        VTDeleteDocumentVariable doc, VTWordImageScaleVariableName(formulaId)
        VTDeleteDocumentVariable doc, VTWordImageCanvasVariableName(formulaId)
        VTDeleteDocumentVariable doc, VTWordImageInkCenterVariableName(formulaId)
        VTDeleteDocumentVariable doc, VTWordImageLetterFontVariableName(formulaId)
        VTDeleteDocumentVariable doc, VTWordNativeSignatureVariableName(formulaId)
End Sub

' --- equation layout identical to Windows (WordFormulaHostLayout) -------------

' A display equation: Word's display type, ApplyOmmlLocalTypography and
' ConfigureDisplayParagraph exactly as VisualTeX for Windows applies them.
Public Function VTV2FinishDisplayEquation( _
    ByVal equationRange As Range, ByVal fontSizePt As Double) As Range

    Dim nativeEquation As OMath
    Dim exact As Range

    If equationRange Is Nothing Or equationRange.OMaths.Count < 1 Then
        Err.Raise vbObjectError + 7473, "VisualTeX", _
            "VisualTeX cannot lay out a missing display equation."
    End If
    Set nativeEquation = VTNativeMathRoot(equationRange.OMaths(1))
    If nativeEquation.Type <> wdOMathDisplay Then nativeEquation.Type = wdOMathDisplay
    Set exact = nativeEquation.Range.Duplicate
    VTV2ApplyEquationTypography exact, fontSizePt
    VTV2ConfigureDisplayParagraph exact
    Set VTV2FinishDisplayEquation = VTNativeMathRoot(exact.OMaths(1)).Range.Duplicate
End Function

' ConfigureDisplayParagraph: the paragraph of a display equation.
Public Sub VTV2ConfigureDisplayParagraph(ByVal equationRange As Range)
    With equationRange.Paragraphs(1).Range.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .SpaceBefore = 0!
        .SpaceAfter = 0!
        .LeftIndent = 0!
        .RightIndent = 0!
        .FirstLineIndent = 0!
    End With
End Sub

' Word shows an inline equation that is alone in its paragraph as a display
' equation. VisualTeX for Windows then lays that paragraph out as a display
' paragraph (WordOmmlHostWriter.Insert: a display host gets
' ConfigureDisplayParagraph).
Public Sub VTV2ConfigurePromotedInlineEquation(ByVal equationRange As Range)
    If equationRange Is Nothing Then Exit Sub
    If equationRange.OMaths.Count < 1 Then Exit Sub
    If VTNativeMathRoot(equationRange.OMaths(1)).Type = wdOMathDisplay Then _
        VTV2ConfigureDisplayParagraph equationRange
End Sub

Public Function VTV2EquationInventory(ByVal doc As Document) As String
    VTV2EquationInventory = CStr(doc.OMaths.Count) & "|" & CStr(doc.InlineShapes.Count)
End Function
