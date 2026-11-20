Attribute VB_Name = "VTWordDisplayNumbering"
Option Explicit
Option Private Module

' =====================================================================
' Batch numbering of display formulas (Windows' Batch Numbering:
' ThisAddIn.DisplayNumbering and WordFormulaService.DisplayNumberingCore)
'
' Numbers every unnumbered display formula of the selection or the
' document: VisualTeX image formulas with image numbers and Word equations
' with equation format v2 numbers, each family in its own series and in the
' document's number format. Numbered formulas keep their numbers; the
' numbers after the new ones follow on and cross-references are updated.
' Formulas in tables, in paragraphs with other content or in an older
' VisualTeX form are left alone and counted.
'
' Performance contract: every document-wide step runs once per command,
' never once per formula. One read-only pass finds the formulas
' (paragraph-local equations, never Document.OMaths; one InlineShapes
' walk); each family's numbering context (existing numbers, headings,
' prefixes, restarts) is planned once; one app request numbers every
' equation and marks every image's metadata numbered; one Undo record
' writes each formula locally, from the end of the document backwards;
' then each family's numbers are updated once and the cross-references
' once. The single-formula planners (VTV2PlanInsertion,
' VTEquationStartsNumberScope, VTReconcileNumberingAfterEdit) never run per
' formula here.
' =====================================================================

' Batch numbering (VTNumberDisplayFormulas): the unattended runner's sentinel
' and result file, and the size of one request to the app.
Private Const VT_WORD_DISPLAY_NUMBERING_UNATTENDED_SENTINEL As String = _
    "/word-display-numbering-unattended.enabled"
Private Const VT_WORD_DISPLAY_NUMBERING_RESULT_FILE As String = _
    "/word-display-numbering-result.txt"
Private Const VT_WORD_DISPLAY_NUMBERING_MAX_REQUEST_BYTES As Long = 1572864
Private Const VT_WORD_DISPLAY_NUMBERING_MAX_REQUEST_ITEMS As Long = 1000
Private Const VT_WORD_DISPLAY_NUMBERING_PROBES As Long = 10
Private VT_WORD_DISPLAY_NUMBERING_WROTE As Boolean

Public Sub VTNumberDisplayFormulas(ByVal wholeDocument As Boolean)
    Dim doc As Document
    Dim unattended As Boolean
    Dim startedAt As Single
    Dim equations As Collection
    Dim images As Collection
    Dim numberedCount As Long
    Dim tableCount As Long
    Dim otherCount As Long
    Dim olderCount As Long
    Dim scopeText As String
    Dim details As String
    Dim message As String
    Dim imagePrefixes() As String
    Dim imageRestarts() As Boolean
    Dim imageTemplates() As Range
    Dim imageValues() As Long
    Dim templateDocument As Document
    Dim restartLosses As Collection
    Dim equationItems() As Variant
    Dim equationSeries() As String
    Dim transfers As Collection
    Dim headingMarks As Collection
    Dim requests As Collection
    Dim equationSources() As Range
    Dim equationErrors() As String
    Dim imageMetadata() As String
    Dim imageErrors() As String
    Dim stagingDocuments As Collection
    Dim staging As Variant
    Dim previousScreenUpdating As Boolean
    Dim screenUpdatingCaptured As Boolean
    Dim mutationStarted As Boolean
    Dim scopeStarted As Boolean
    Dim recordStarted As Boolean
    Dim equationsWritten As Long
    Dim imagesWritten As Long
    Dim failedCount As Long
    Dim firstTarget As Long
    Dim request As Variant
    Dim operationStage As String
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim uniformWidth As Single

    If VT_WORD_NUMBER_UPDATE_RUNNING Then Exit Sub
    On Error Resume Next
    unattended = VTPathFileExists( _
        VTApplicationSupportRoot() & VT_WORD_DISPLAY_NUMBERING_UNATTENDED_SENTINEL)
    Err.Clear
    On Error GoTo Failed
    operationStage = "document"
    VTRequireWritableWordDocument
    Set doc = ActiveDocument
    If doc.TrackRevisions Then
        Err.Raise vbObjectError + 7620, "VisualTeX", _
            "Turn off Track Changes before numbering display equations."
    End If
    If Not wholeDocument Then
        If Selection.StoryType <> wdMainTextStory Then
            Err.Raise vbObjectError + 7620, "VisualTeX", _
                "Select display equations in the body of the document."
        End If
    End If
    startedAt = Timer
    VTWordPerformanceStartLocal "display-numbering"
    VT_WORD_NUMBER_UPDATE_RUNNING = True
    VTBeginNumberingResolveScope
    scopeStarted = True
    Application.StatusBar = "VisualTeX: finding display equations..."

    operationStage = "plan"
    VTPlanDisplayNumbering doc, wholeDocument, equations, images, _
        numberedCount, tableCount, otherCount, olderCount
    VTWordPerformanceNote "equations", CStr(equations.Count)
    VTWordPerformanceNote "images", CStr(images.Count)
    VTWordPerformanceMark "planned"
    If wholeDocument Then scopeText = "the document" Else scopeText = "the selection"
    details = VTDisplayNumberingDetails( _
        numberedCount, tableCount, otherCount, olderCount)
    If equations.Count + images.Count = 0 Then
        Application.StatusBar = "VisualTeX: no display equation needed a number."
        VTFinishDisplayNumberingReport unattended, "nothing", _
            "There is no unnumbered display equation in " & scopeText & "." & details
        GoTo Finished
    End If
    If Not unattended And Not VT_AGENT_MODE Then
        message = "Add equation numbers to " & _
            CStr(equations.Count + images.Count) & _
            " unnumbered display equation(s) in " & scopeText & "?" & vbCr & vbCr
        If images.Count > 0 Then message = message & _
            "VisualTeX image formulas: " & CStr(images.Count) & vbCr
        If equations.Count > 0 Then message = message & _
            "Word OMML equations: " & CStr(equations.Count) & vbCr
        message = message & "Number format: " & _
            VTDisplayNumberingFormatName(doc) & "." & vbCr & vbCr & _
            "Existing numbers stay as they are; the numbers after the new ones " & _
            "follow on and cross-references are updated." & details & vbCr & vbCr & _
            "Undo reverses the whole operation."
        If MsgBox(message, vbOKCancel + vbQuestion, "VisualTeX") <> vbOK Then
            Application.StatusBar = _
                "VisualTeX: batch numbering was cancelled; the document was not changed."
            GoTo Finished
        End If
    End If

    ' ---- Preflight: everything is read and prepared before Undo starts. ----
    previousScreenUpdating = Application.ScreenUpdating
    screenUpdatingCaptured = True
    Application.ScreenUpdating = False
    Application.StatusBar = "VisualTeX: preparing " & _
        CStr(equations.Count + images.Count) & " equation numbers..."
    Set transfers = New Collection
    If images.Count > 0 Then
        operationStage = "plan-image-numbers"
        VTPlanDisplayNumberingImageNumbers doc, images, imagePrefixes, _
            imageRestarts, imageValues, restartLosses
        imageTemplates = VTBuildImageNumberTemplates( _
            doc, imagePrefixes, imageRestarts, imageValues, templateDocument)
        VTWordPerformanceMark "image-numbers-planned"
    End If
    If equations.Count > 0 Then
        operationStage = "plan-equation-numbers"
        VTPlanDisplayNumberingEquationItems doc, equations, equationItems, _
            transfers, headingMarks
        equationSeries = VTDisplayNumberingSeries(equationItems)
        VTWordPerformanceMark "equation-numbers-planned"
    End If
    operationStage = "request"
    Set requests = VTRequestDisplayNumbering( _
        doc, equations, equationItems, images, _
        equationSources, equationErrors, imageMetadata, imageErrors, _
        stagingDocuments)
    doc.Activate
    VTWordPerformanceMark "requested"

    ' ---- One Undo record: every write, then one renumbering pass. ----
    operationStage = "write"
    VTBeginWordInternalMutation
    mutationStarted = True
    VT_WORD_DISPLAY_NUMBERING_WROTE = False
    Application.UndoRecord.StartCustomRecord "VisualTeX Number Display Equations"
    recordStarted = True
    If images.Count > 0 Then
        Set VT_WORD_DISPLAY_NUMBERING_IMAGE_STYLE = Nothing
        Set VT_WORD_DISPLAY_NUMBERING_IMAGE_STYLE = VTEnsureNumberedImageParagraphStyle(doc)
        VT_WORD_DISPLAY_NUMBERING_LAYOUT_WIDTH = 0!
        ' One width for the whole batch only where every formula has it: one
        ' section with one column or equal columns. With unequal columns each
        ' formula takes the column it is in.
        If doc.Sections.Count = 1 Then
            If VTEquationUniformColumnWidth(doc.Sections(1).PageSetup, uniformWidth) Then _
                VT_WORD_DISPLAY_NUMBERING_LAYOUT_WIDTH = uniformWidth
        End If
    End If
    VTWriteDisplayNumbering doc, equations, images, equationSources, _
        equationErrors, imageMetadata, imageErrors, imagePrefixes, _
        imageRestarts, imageValues, imageTemplates, equationsWritten, _
        imagesWritten, failedCount, firstTarget
    VTWordPerformanceMark "written"
    operationStage = "renumber"
    Application.StatusBar = "VisualTeX: updating equation numbers..."
    VTRenumberDisplayNumbering doc, imagesWritten, equationsWritten, _
        transfers, restartLosses, equationSeries, headingMarks, firstTarget
    VTWordPerformanceMark "renumbered"
    Application.UndoRecord.EndCustomRecord
    recordStarted = False
    Application.StatusBar = "VisualTeX: numbered " & _
        CStr(equationsWritten + imagesWritten) & " display equation(s) in " & _
        Format$(VTTimerElapsedSeconds(startedAt, Timer), "0.0") & " s."
    message = "Numbered " & CStr(equationsWritten + imagesWritten) & _
        " display equation(s): " & CStr(imagesWritten) & " image formula(s), " & _
        CStr(equationsWritten) & " OMML equation(s)."
    If failedCount > 0 Then
        message = message & vbCr & CStr(failedCount) & _
            " equation(s) could not be read and were left unnumbered."
    End If
    VTFinishDisplayNumberingReport unattended, "complete", message & details, _
        "equationsWritten=" & CStr(equationsWritten) & vbLf & _
        "imagesWritten=" & CStr(imagesWritten) & vbLf & _
        "failed=" & CStr(failedCount) & vbLf & _
        "alreadyNumbered=" & CStr(numberedCount) & vbLf & _
        "inTable=" & CStr(tableCount) & vbLf & _
        "otherContent=" & CStr(otherCount) & vbLf & _
        "olderFormat=" & CStr(olderCount) & vbLf & _
        "seconds=" & Format$(VTTimerElapsedSeconds(startedAt, Timer), "0.00"), _
        Not unattended And failedCount > 0

Finished:
    On Error Resume Next
    Set VT_WORD_DISPLAY_NUMBERING_IMAGE_STYLE = Nothing
    VT_WORD_DISPLAY_NUMBERING_LAYOUT_WIDTH = 0!
    If Not stagingDocuments Is Nothing Then
        For Each staging In stagingDocuments
            staging.Close SaveChanges:=wdDoNotSaveChanges
        Next staging
    End If
    If Not templateDocument Is Nothing Then templateDocument.Close SaveChanges:=wdDoNotSaveChanges
    If Not doc Is Nothing Then doc.Activate
    If Not requests Is Nothing Then
        For Each request In requests
            VTDeleteSessionFiles CStr(request)
        Next request
    End If
    If mutationStarted Then VTEndWordInternalMutation
    If screenUpdatingCaptured Then Application.ScreenUpdating = previousScreenUpdating
    If scopeStarted Then VTEndNumberingResolveScope
    VT_WORD_NUMBER_UPDATE_RUNNING = False
    If VT_WORD_PERFORMANCE_ACTIVE Then
        VTWordPerformanceFlush VTNewUuidV4(), failureNumber, failureDescription
    End If
    Err.Clear
    On Error GoTo 0
    If failureNumber <> 0 Then
        Application.StatusBar = "VisualTeX: batch numbering failed; the document was restored."
        If unattended Then
            VTFinishDisplayNumberingReport True, "failed", failureDescription, _
                "error=" & CStr(failureNumber) & vbLf & "stage=" & operationStage
        Else
            VTShowError "batch numbering", failureNumber, failureDescription
        End If
    End If
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    If failureNumber = 0 Then failureNumber = vbObjectError + 7620
    On Error Resume Next
    If recordStarted Then
        Application.UndoRecord.EndCustomRecord
        recordStarted = False
    End If
    ' Undo only a record that holds a change: an empty record leaves no Undo
    ' step, and Undo would then revert the user's own last action.
    If VT_WORD_DISPLAY_NUMBERING_WROTE And Not doc Is Nothing Then doc.Undo 1
    VT_WORD_DISPLAY_NUMBERING_WROTE = False
    Resume Finished
End Sub

' The unattended runner (sentinel VT_WORD_DISPLAY_NUMBERING_UNATTENDED_SENTINEL)
' reads the result from a file; a person reads it in a message.
Private Sub VTFinishDisplayNumberingReport( _
    ByVal unattended As Boolean, _
    ByVal status As String, _
    ByVal message As String, _
    Optional ByVal resultRows As String = "", _
    Optional ByVal showMessage As Boolean = True)

    If unattended Then
        On Error Resume Next
        VTWriteTextAtomic _
            VTApplicationSupportRoot() & VT_WORD_DISPLAY_NUMBERING_RESULT_FILE, _
            "status=" & status & vbLf & resultRows & vbLf & _
            "message=" & Replace$(Replace$(message, vbCr, " | "), vbLf, " | ") & vbLf
        Err.Clear
        On Error GoTo 0
    ElseIf showMessage Then
        VTShowInformation message
    End If
End Sub

Private Function VTDisplayNumberingDetails( _
    ByVal numberedCount As Long, _
    ByVal tableCount As Long, _
    ByVal otherCount As Long, _
    ByVal olderCount As Long) As String

    Dim lines As String

    If numberedCount > 0 Then lines = lines & vbCr & _
        "Already numbered (unchanged): " & CStr(numberedCount)
    If tableCount > 0 Then lines = lines & vbCr & _
        "Inside a table, skipped: " & CStr(tableCount)
    If otherCount > 0 Then lines = lines & vbCr & _
        "The paragraph holds other content, skipped: " & CStr(otherCount)
    If olderCount > 0 Then lines = lines & vbCr & _
        "Older VisualTeX format, skipped: " & CStr(olderCount) & _
        " (run Update Equation Format first)"
    If Len(lines) > 0 Then lines = vbCr & lines
    VTDisplayNumberingDetails = lines
End Function

Private Function VTDisplayNumberingFormatName(ByVal doc As Document) As String
    Dim numberingMode As String
    Dim separatorText As String

    VTResolveEquationNumbering doc, numberingMode, separatorText
    Select Case numberingMode
        Case VT_WORD_NUMBERING_MODE_CHAPTER
            If separatorText = "-" Then
                VTDisplayNumberingFormatName = "by chapter (2-1)"
            Else
                VTDisplayNumberingFormatName = "by chapter (2.1)"
            End If
        Case VT_WORD_NUMBERING_MODE_SECTION
            If separatorText = "-" Then
                VTDisplayNumberingFormatName = "by section (2.3-1)"
            Else
                VTDisplayNumberingFormatName = "by section (2.3.1)"
            End If
        Case Else
            VTDisplayNumberingFormatName = "sequential (1)"
    End Select
End Function

' --- plan: which formulas get a number ---------------------------------------

' The display formulas of the selection or the document that batch numbering
' numbers, in document order (equations: Array(start, end, font size); image
' formulas: Array(start, formulaId, metadata)), and how many it leaves alone
' and why. Reads only.
Private Sub VTPlanDisplayNumbering( _
    ByVal doc As Document, _
    ByVal wholeDocument As Boolean, _
    ByRef equations As Collection, _
    ByRef images As Collection, _
    ByRef numberedCount As Long, _
    ByRef tableCount As Long, _
    ByRef otherCount As Long, _
    ByRef olderCount As Long)

    Dim scope As Range
    Dim candidateMath As OMath
    Dim rootMath As OMath
    Dim shape As InlineShape
    Dim lastStart As Long

    Set equations = New Collection
    Set images = New Collection
    If wholeDocument Then
        Set scope = doc.Content
    Else
        ' A display formula owns its paragraph: a selection takes in every
        ' paragraph it touches, a caret its own paragraph.
        Set scope = Selection.Range.Duplicate
        scope.Start = scope.Paragraphs(1).Range.Start
        scope.End = scope.Paragraphs.Last.Range.End
    End If
    lastStart = -1
    For Each candidateMath In VTOMathsInDocumentOrder(scope)
        Set rootMath = VTNativeMathRoot(candidateMath)
        If rootMath.Range.Start > lastStart Then
            lastStart = rootMath.Range.Start
            VTPlanDisplayNumberingEquation rootMath, equations, _
                numberedCount, tableCount, otherCount, olderCount
        End If
    Next candidateMath
    For Each shape In scope.InlineShapes
        VTPlanDisplayNumberingImage shape, images, _
            numberedCount, tableCount, otherCount, olderCount
    Next shape
End Sub

Private Sub VTPlanDisplayNumberingEquation( _
    ByVal rootMath As OMath, _
    ByVal equations As Collection, _
    ByRef numberedCount As Long, _
    ByRef tableCount As Long, _
    ByRef otherCount As Long, _
    ByRef olderCount As Long)

    Dim doc As Document
    Dim exact As Range
    Dim paragraphRange As Range
    Dim mark As Bookmark
    Dim fontSizePt As Double

    If rootMath.Type <> wdOMathDisplay Then Exit Sub
    Set exact = rootMath.Range.Duplicate
    Set doc = exact.Document
    If exact.Fields.Count > 0 Then
        ' A field in an equation without a current number is an older
        ' VisualTeX number (REF to an external SEQ helper) or the user's own.
        If VTEquationHasNumberStructure(rootMath, exact) Then
            numberedCount = numberedCount + 1
        Else
            olderCount = olderCount + 1
        End If
        Exit Sub
    End If
    If InStr(1, exact.Text, "#", vbBinaryCompare) > 0 Then
        If VTOMathHasEquationArray(rootMath) Then
            ' A number typed as #(...) in Word's equation editor.
            numberedCount = numberedCount + 1
            Exit Sub
        End If
    End If
    Set paragraphRange = exact.Paragraphs(1).Range.Duplicate
    If VTRangeIsWithinTable(paragraphRange) Then
        tableCount = tableCount + 1
        Exit Sub
    End If
    If exact.InlineShapes.Count > 0 Then
        ' A picture inside the equation, such as an image formula.
        otherCount = otherCount + 1
        Exit Sub
    End If
    For Each mark In exact.Bookmarks
        If Left$(mark.Name, 3) = "VT_" Then
            ' Older VisualTeX tags; Update Equation Format retires them.
            olderCount = olderCount + 1
            Exit Sub
        End If
    Next mark
    If paragraphRange.OMaths.Count <> 1 Or _
       paragraphRange.InlineShapes.Count <> 0 Or _
       VTWordRangeHasMeaningfulText( _
           doc.Range(paragraphRange.Start, exact.Start)) Or _
       VTWordRangeHasMeaningfulText( _
           doc.Range(exact.End, paragraphRange.End)) Then
        otherCount = otherCount + 1
        Exit Sub
    End If
    fontSizePt = exact.Font.Size
    If fontSizePt < 5# Or fontSizePt > 200# Then fontSizePt = 0#
    equations.Add Array(exact.Start, exact.End, fontSizePt)
End Sub

Private Function VTOMathHasEquationArray(ByVal rootMath As OMath) As Boolean
    Dim mathFunction As OMathFunction

    For Each mathFunction In rootMath.Functions
        If mathFunction.Type = wdOMathFunctionEqArray Then
            VTOMathHasEquationArray = True
            Exit Function
        End If
    Next mathFunction
End Function

' A native equation number as batch numbering writes it and Word or VisualTeX
' numbers an equation: a SEQ field in an equation array. VTPureNativeNumberField
' also finds the # marker with Find, which costs about 9 ms an equation, and
' several times more once the Word session has shown an image formula.
Private Function VTEquationHasNumberStructure( _
    ByVal rootMath As OMath, ByVal exact As Range) As Boolean

    Dim candidate As Field

    If Not VTOMathHasEquationArray(rootMath) Then Exit Function
    For Each candidate In exact.Fields
        If candidate.Type = wdFieldSequence Then
            VTEquationHasNumberStructure = True
            Exit Function
        End If
    Next candidate
End Function

Private Sub VTPlanDisplayNumberingImage( _
    ByVal shape As InlineShape, _
    ByVal images As Collection, _
    ByRef numberedCount As Long, _
    ByRef tableCount As Long, _
    ByRef otherCount As Long, _
    ByRef olderCount As Long)

    Dim doc As Document
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim metadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean
    Dim paragraphRange As Range
    Dim container As Range

    ' Ribbon-style query: never repairs the shape while planning.
    If Not VTTryResolveVisualTeXInlineShapeReference( _
       shape, formulaId, displayMode, numbered, metadata, _
       metadataNeedsWrite, formatNeedsWrite, False) Then Exit Sub
    If displayMode <> "block" Then Exit Sub
    ' An image inside an equation: VTPlanDisplayNumberingEquation counts it.
    If shape.Range.OMaths.Count > 0 Then Exit Sub
    If numbered Then
        numberedCount = numberedCount + 1
        Exit Sub
    End If
    If Not VTImagePlaceRef(shape.Range) Is Nothing Then
        ' A paragraph number that the formula's own state does not record.
        numberedCount = numberedCount + 1
        Exit Sub
    End If
    Set doc = shape.Range.Document
    Set paragraphRange = shape.Range.Paragraphs(1).Range.Duplicate
    If VTRangeIsWithinTable(paragraphRange) Then
        tableCount = tableCount + 1
        Exit Sub
    End If
    If Not VTIsEncodedMetadata(metadata) Then
        olderCount = olderCount + 1
        Exit Sub
    End If
    Set container = VTVisualTeXImageContainerRange(shape)
    If paragraphRange.InlineShapes.Count <> 1 Or _
       VTRangeHasMeaningfulMath(paragraphRange) Or _
       VTWordRangeHasMeaningfulText( _
           doc.Range(paragraphRange.Start, container.Start)) Or _
       VTWordRangeHasMeaningfulText( _
           doc.Range(container.End, paragraphRange.End)) Then
        otherCount = otherCount + 1
        Exit Sub
    End If
    images.Add Array(shape.Range.Start, formulaId, metadata)
End Sub

' --- plan: the numbers --------------------------------------------------------

' The chapter prefix and restart of every new image number, by the rule of
' VTUpgradeWordNumberFormats: the new numbers and the image numbers already
' in the document form one series, numbered in one heading pass. losses holds
' (as Ranges that follow the edits before them) the numbered images that start
' their chapter now but follow a new number afterwards, and so lose their \r 1.
' values holds the number each new image shows.
Private Sub VTPlanDisplayNumberingImageNumbers( _
    ByVal doc As Document, _
    ByVal images As Collection, _
    ByRef prefixes() As String, _
    ByRef restarts() As Boolean, _
    ByRef values() As Long, _
    ByRef losses As Collection)

    Dim numberingMode As String
    Dim existing() As Long
    Dim existingCount As Long
    Dim anchors() As Long
    Dim anchorCount As Long
    Dim targetAnchors() As Long
    Dim anchorPrefixes() As String
    Dim anchorOrdinals() As Long
    Dim shape As InlineShape
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim target As Variant
    Dim existingIndex As Long
    Dim targetIndex As Long
    Dim existingSeen As Boolean
    Dim isTarget As Boolean

    Set losses = New Collection
    ReDim prefixes(1 To images.Count)
    ReDim restarts(1 To images.Count)
    ReDim values(1 To images.Count)
    numberingMode = VTEquationNumberingMode(doc)

    ReDim existing(1 To 16)
    For Each shape In doc.InlineShapes
        If VTTryParseFormulaReference( _
           shape.Title, formulaId, displayMode, numbered) Then
            If numbered And displayMode = "block" Then
                existingCount = existingCount + 1
                If existingCount > UBound(existing) Then _
                    ReDim Preserve existing(1 To 2 * UBound(existing))
                existing(existingCount) = shape.Range.Start
            End If
        End If
    Next shape

    ' Both lists are in document order: merge them.
    ReDim anchors(1 To existingCount + images.Count)
    ReDim targetAnchors(1 To images.Count)
    existingIndex = 1
    targetIndex = 1
    For Each target In images
        Do While existingIndex <= existingCount
            If existing(existingIndex) > target(0) Then Exit Do
            anchorCount = anchorCount + 1
            anchors(anchorCount) = existing(existingIndex)
            existingIndex = existingIndex + 1
        Loop
        anchorCount = anchorCount + 1
        anchors(anchorCount) = target(0)
        targetAnchors(targetIndex) = anchorCount
        targetIndex = targetIndex + 1
    Next target
    Do While existingIndex <= existingCount
        anchorCount = anchorCount + 1
        anchors(anchorCount) = existing(existingIndex)
        existingIndex = existingIndex + 1
    Loop

    If VTEquationNumberingRestartLevel(doc) < 1 Then
        ' One series through the document: a number counts every one before it.
        For targetIndex = 1 To images.Count
            values(targetIndex) = targetAnchors(targetIndex)
        Next targetIndex
        Exit Sub
    End If
    VTBuildEquationHeadingPrefixes doc, anchors, anchorCount, numberingMode, _
        anchorPrefixes, anchorOrdinals
    For targetIndex = 1 To images.Count
        prefixes(targetIndex) = anchorPrefixes(targetAnchors(targetIndex))
        restarts(targetIndex) = (anchorOrdinals(targetAnchors(targetIndex)) = 1)
        values(targetIndex) = anchorOrdinals(targetAnchors(targetIndex))
    Next targetIndex

    ' An existing number that was the first of its chapter (no existing
    ' number before it there) but is no longer the first one loses its restart.
    targetIndex = 1
    For existingIndex = 1 To anchorCount
        If anchorOrdinals(existingIndex) = 1 Then existingSeen = False
        isTarget = False
        If targetIndex <= images.Count Then _
            isTarget = (targetAnchors(targetIndex) = existingIndex)
        If isTarget Then
            targetIndex = targetIndex + 1
        Else
            If Not existingSeen And anchorOrdinals(existingIndex) > 1 Then _
                losses.Add doc.Range(anchors(existingIndex), anchors(existingIndex) + 1)
            existingSeen = True
        End If
    Next existingIndex
End Sub

' The number of every new equation (Windows WordNativeNumberingAdoptionPlan.
' PrepareBatchInsertion), from one read of the existing v2 numbers and the
' headings: a new number joins the series of the nearest number before it
' (the first series when none is before it), and only the earliest new number
' before a series' first number takes over that number's explicit start
' (\r N), which the first number then loses (transfers). items(i) is
' equation i's number: Array(SEQ label, heading level, separator, literal
' prefix, automatic prefix, restart, initial result).
Private Sub VTPlanDisplayNumberingEquationItems( _
    ByVal doc As Document, _
    ByVal equations As Collection, _
    ByRef items() As Variant, _
    ByVal transfers As Collection, _
    ByRef headingMarks As Collection)

    Dim entryList As Collection
    Dim entries() As Variant
    Dim entryCount As Long
    Dim entry As Variant
    Dim headings As Collection
    Dim headingLevel As Long
    Dim separator As String
    Dim defaultName As String
    Dim target As Variant
    Dim index As Long
    Dim nearest As Long
    Dim first As Long
    Dim position As Long
    Dim sequenceName As String
    Dim literalPrefix As String
    Dim automatic As Boolean
    Dim scopeStart As Long
    Dim numberText As String
    Dim restart As Long
    Dim initial As Long
    Dim takenRuns As Collection
    Dim counters As Collection
    Dim counterKey As String

    VTV2DocumentFormat doc, headingLevel, separator
    Set entryList = VTV2NumberedEquations(doc)
    entryCount = entryList.Count
    If entryCount > 0 Then
        ReDim entries(1 To entryCount)
        For Each entry In entryList
            index = index + 1
            entries(index) = entry
        Next entry
    Else
        defaultName = VTNativeEquationLabelName()
    End If
    If headingLevel > 0 Then Set headings = VTV2CollectHeadings(doc, headingLevel)
    ' The renumbering pass counts \s restarts at these headings: Ranges follow
    ' the writes, and searching the edited document for its headings again
    ' costs many times more on Word for Mac.
    Set headingMarks = New Collection
    If Not headings Is Nothing Then
        For Each entry In headings
            headingMarks.Add Array(doc.Range(entry(0), entry(0)), entry(1))
        Next entry
    End If
    Set takenRuns = New Collection
    Set counters = New Collection
    ReDim items(1 To equations.Count)
    index = 0
    For Each target In equations
        index = index + 1
        position = target(0)
        restart = 0
        ' Existing numbers before this one set the series counters (the
        ' initial result; the renumbering pass recounts every field).
        Do While nearest < entryCount
            If entries(nearest + 1)(0) >= position Then Exit Do
            nearest = nearest + 1
            VTDisplayNumberingCount counters, _
                VTDisplayNumberingCounterKey(entries(nearest)(1), headings, _
                    entries(nearest)(0), headingLevel), entries(nearest)(5)
        Loop
        If entryCount = 0 Then
            sequenceName = defaultName
        Else
            first = nearest
            If first = 0 Then first = 1
            Do While first > 1
                If entries(first - 1)(3) <> entries(first)(3) Or _
                   StrComp(entries(first - 1)(1), entries(first)(1), _
                       vbTextCompare) <> 0 Then Exit Do
                first = first - 1
            Loop
            sequenceName = entries(first)(1)
            If Not VTCollectionHasKey(takenRuns, CStr(first)) Then
                takenRuns.Add first, CStr(first)
                If position <= entries(first)(0) And entries(first)(6) > 0 Then
                    restart = entries(first)(6)
                    transfers.Add doc.Range(entries(first)(0), entries(first)(8))
                End If
            End If
        End If
        literalPrefix = ""
        automatic = False
        If headingLevel > 0 Then
            numberText = VTV2HeadingScope(headings, position, headingLevel, _
                scopeStart, automatic)
            If scopeStart < 0 Then automatic = False
            If Not automatic Then literalPrefix = numberText
        End If
        counterKey = VTDisplayNumberingCounterKey(sequenceName, headings, _
            position, headingLevel)
        If restart > 0 Then
            initial = restart
        Else
            initial = 1
            If VTCollectionHasKey(counters, counterKey) Then _
                initial = CLng(counters(counterKey)) + 1
        End If
        VTDisplayNumberingCount counters, counterKey, initial
        items(index) = Array(sequenceName, headingLevel, separator, _
            literalPrefix, automatic, restart, initial)
    Next target
End Sub

' The series counter of a number at position: its SEQ label, and in chapter or
' section numbering the heading scope it restarts in.
Private Function VTDisplayNumberingCounterKey( _
    ByVal sequenceName As String, _
    ByVal headings As Collection, _
    ByVal position As Long, _
    ByVal headingLevel As Long) As String

    Dim scopeStart As Long
    Dim automatic As Boolean

    VTDisplayNumberingCounterKey = LCase$(sequenceName)
    If headingLevel > 0 Then
        Call VTV2HeadingScope(headings, position, headingLevel, scopeStart, automatic)
        VTDisplayNumberingCounterKey = VTDisplayNumberingCounterKey & "|" & CStr(scopeStart)
    End If
End Function

Private Sub VTDisplayNumberingCount( _
    ByVal counters As Collection, _
    ByVal counterKey As String, _
    ByVal value As Long)

    If VTCollectionHasKey(counters, counterKey) Then counters.Remove counterKey
    counters.Add value, counterKey
End Sub

' --- the app request ----------------------------------------------------------

' Word's export of one equation for the request: its w:document start tag
' (the namespace declarations) and its m:oMath, escaped for one manifest row.
Private Sub VTDisplayNumberingExportEquation( _
    ByVal exact As Range, _
    ByRef rootTag As String, _
    ByRef mathXml As String)

    Dim xml As String
    Dim rootEnd As Long
    Dim mathStart As Long
    Dim mathEnd As Long

    xml = VTFormulaRestoreNativeXml(exact)
    If Left$(xml, 11) <> "<w:document" Then
        Err.Raise vbObjectError + 7621, "VisualTeX", _
            "Word did not export an equation to number."
    End If
    rootEnd = InStr(1, xml, ">", vbBinaryCompare)
    rootTag = Left$(xml, rootEnd)
    mathStart = InStr(rootEnd, xml, "<m:oMath>", vbBinaryCompare)
    mathEnd = InStrRev(xml, "</m:oMath>", -1, vbBinaryCompare)
    If mathStart = 0 Or mathEnd < mathStart Then
        Err.Raise vbObjectError + 7621, "VisualTeX", _
            "Word did not export an equation to number."
    End If
    mathXml = Mid$(xml, mathStart, mathEnd + Len("</m:oMath>") - mathStart)
    mathXml = Replace$(Replace$(Replace$(mathXml, "%", "%25"), vbCr, "%0D"), vbLf, "%0A")
End Sub

' Sends the equations and image metadata to the app and opens the numbered
' equations: one CLI run per request of at most
' VT_WORD_DISPLAY_NUMBERING_MAX_REQUEST_BYTES characters and
' VT_WORD_DISPLAY_NUMBERING_MAX_REQUEST_ITEMS formulas of a kind.
' sources(i) is equation i's numbered form in its staging package, or Nothing
' with the reason in equationErrors(i); imageMetadata(i) is image i's
' numbered metadata, or empty with the reason in imageErrors(i). Returns the
' request session ids for cleanup.
Private Function VTRequestDisplayNumbering( _
    ByVal doc As Document, _
    ByVal equations As Collection, _
    ByRef equationItems() As Variant, _
    ByVal images As Collection, _
    ByRef sources() As Range, _
    ByRef equationErrors() As String, _
    ByRef imageMetadata() As String, _
    ByRef imageErrors() As String, _
    ByRef stagingDocuments As Collection) As Collection

    Dim requests As New Collection
    Dim rows() As String
    Dim rowCount As Long
    Dim rowBytes As Long
    Dim requestEquations() As Long
    Dim requestImages() As Long
    Dim equationCount As Long
    Dim imageCount As Long
    Dim rootTag As String
    Dim itemRoot As String
    Dim mathXml As String
    Dim row As String
    Dim target As Variant
    Dim index As Long
    Dim exportErrorNumber As Long
    Dim exportErrorDescription As String

    Set stagingDocuments = New Collection
    If equations.Count > 0 Then
        ReDim sources(1 To equations.Count)
        ReDim equationErrors(1 To equations.Count)
        ReDim requestEquations(1 To equations.Count)
    End If
    If images.Count > 0 Then
        ReDim imageMetadata(1 To images.Count)
        ReDim imageErrors(1 To images.Count)
        ReDim requestImages(1 To images.Count)
    End If
    ReDim rows(1 To 64)

    index = 0
    For Each target In equations
        index = index + 1
        If index Mod 10 = 1 Then
            Application.StatusBar = "VisualTeX: reading equation " & CStr(index) & _
                " of " & CStr(equations.Count) & "..."
        End If
        On Error Resume Next
        VTDisplayNumberingExportEquation doc.Range(target(0), target(1)), itemRoot, mathXml
        exportErrorNumber = Err.Number
        exportErrorDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If exportErrorNumber <> 0 Then
            equationErrors(index) = exportErrorDescription
        Else
            If Len(rootTag) = 0 Then rootTag = itemRoot
            row = VTDisplayNumberingEquationRow( _
                equationCount + 1, equationItems(index), mathXml, itemRoot, rootTag)
            If equationCount > 0 And _
               (rowBytes + Len(row) > VT_WORD_DISPLAY_NUMBERING_MAX_REQUEST_BYTES Or _
                equationCount >= VT_WORD_DISPLAY_NUMBERING_MAX_REQUEST_ITEMS) Then
                VTSendDisplayNumberingRequest doc, requests, rows, rowCount, rootTag, _
                    requestEquations, equationCount, requestImages, 0, sources, _
                    equationErrors, imageMetadata, imageErrors, stagingDocuments, images
                equationCount = 0
                rowCount = 0
                rowBytes = 0
                row = VTDisplayNumberingEquationRow( _
                    1, equationItems(index), mathXml, itemRoot, rootTag)
            End If
            equationCount = equationCount + 1
            requestEquations(equationCount) = index
            VTDisplayNumberingAddRow rows, rowCount, row
            rowBytes = rowBytes + Len(row)
        End If
    Next target

    index = 0
    For Each target In images
        index = index + 1
        row = "image" & CStr(imageCount + 1) & "Metadata=" & target(2) & vbLf
        If (equationCount + imageCount) > 0 And _
           (rowBytes + Len(row) > VT_WORD_DISPLAY_NUMBERING_MAX_REQUEST_BYTES Or _
            imageCount >= VT_WORD_DISPLAY_NUMBERING_MAX_REQUEST_ITEMS) Then
            VTSendDisplayNumberingRequest doc, requests, rows, rowCount, rootTag, _
                requestEquations, equationCount, requestImages, imageCount, sources, _
                equationErrors, imageMetadata, imageErrors, stagingDocuments, images
            equationCount = 0
            imageCount = 0
            rowCount = 0
            rowBytes = 0
            row = "image1Metadata=" & target(2) & vbLf
        End If
        imageCount = imageCount + 1
        requestImages(imageCount) = index
        VTDisplayNumberingAddRow rows, rowCount, row
        rowBytes = rowBytes + Len(row)
    Next target
    If equationCount + imageCount > 0 Then
        VTSendDisplayNumberingRequest doc, requests, rows, rowCount, rootTag, _
            requestEquations, equationCount, requestImages, imageCount, sources, _
            equationErrors, imageMetadata, imageErrors, stagingDocuments, images
    End If
    Set VTRequestDisplayNumbering = requests
End Function

' The request rows of one equation: its number (item = Array(SEQ label,
' heading level, separator, literal prefix, automatic prefix, restart,
' initial result)) and Word's export of it.
Private Function VTDisplayNumberingEquationRow( _
    ByVal itemIndex As Long, _
    ByVal item As Variant, _
    ByVal mathXml As String, _
    ByVal itemRoot As String, _
    ByVal rootTag As String) As String

    VTDisplayNumberingEquationRow = VTV2NumberItemManifest(itemIndex, _
        CStr(item(0)), CLng(item(1)), CStr(item(2)), CStr(item(3)), _
        CBool(item(4)), CLng(item(5)), CLng(item(6))) & _
        "item" & CStr(itemIndex) & "Omml=" & mathXml & vbLf
    If StrComp(itemRoot, rootTag, vbBinaryCompare) <> 0 Then
        VTDisplayNumberingEquationRow = VTDisplayNumberingEquationRow & _
            "item" & CStr(itemIndex) & "OmmlRoot=" & _
            VTBase64UrlEncodeUtf8(itemRoot) & vbLf
    End If
End Function

Private Sub VTDisplayNumberingAddRow( _
    ByRef rows() As String, _
    ByRef rowCount As Long, _
    ByVal row As String)

    rowCount = rowCount + 1
    If rowCount > UBound(rows) Then ReDim Preserve rows(1 To 2 * UBound(rows))
    rows(rowCount) = row
End Sub

' One displayNumbering request. requestEquations(1..equationCount) and
' requestImages(1..imageCount) are the batch indices of its formulas, in the
' order of its rows.
Private Sub VTSendDisplayNumberingRequest( _
    ByVal doc As Document, _
    ByVal requests As Collection, _
    ByRef rows() As String, _
    ByVal rowCount As Long, _
    ByVal rootTag As String, _
    ByRef requestEquations() As Long, _
    ByVal equationCount As Long, _
    ByRef requestImages() As Long, _
    ByVal imageCount As Long, _
    ByRef sources() As Range, _
    ByRef equationErrors() As String, _
    ByRef imageMetadata() As String, _
    ByRef imageErrors() As String, _
    ByVal stagingDocuments As Collection, _
    ByVal images As Collection)

    Dim sessionId As String
    Dim body() As String
    Dim manifest As String
    Dim timingDetail As String
    Dim separatorPosition As Long
    Dim result As Object
    Dim imageResult As Object
    Dim docxPath As String
    Dim imagesPath As String
    Dim staging As Document
    Dim stagingParagraph As Paragraph
    Dim index As Long
    Dim key As String

    sessionId = VTNewUuidV4()
    requests.Add sessionId
    ReDim body(0 To rowCount)
    body(0) = "protocolVersion=" & CStr(VT_PROTOCOL_VERSION) & vbLf & _
        "sessionId=" & sessionId & vbLf & _
        "sourceKind=displayNumbering" & vbLf & _
        "sourceDocumentId=" & VTBase64UrlEncodeUtf8(doc.FullName) & vbLf & _
        "bookmarkName=-" & vbLf & _
        "itemCount=" & CStr(equationCount) & vbLf & _
        "imageCount=" & CStr(imageCount) & vbLf
    If equationCount > 0 Then
        VTApplyChosenOmmlFontToDocument doc
        body(0) = body(0) & _
            "mathFont=" & VTBase64UrlEncodeUtf8(doc.OMathFontName) & vbLf & _
            "ommlRoot=" & VTBase64UrlEncodeUtf8(rootTag) & vbLf
    End If
    For index = 1 To rowCount
        body(index) = rows(index)
    Next index
    manifest = Join(body, "")
    Application.StatusBar = "VisualTeX: numbering " & _
        CStr(equationCount + imageCount) & " equation(s)..."
    If VTTryWriteOmmlBatchManifestDirect(sessionId, manifest) Then
        timingDetail = VTConvertOmmlBatch(VT_WORD_HOST, sessionId)
    Else
        timingDetail = VTConvertOmmlBatch(VT_WORD_HOST, sessionId, manifest)
    End If
    separatorPosition = InStrRev(timingDetail, "|", -1, vbBinaryCompare)
    If separatorPosition <= 1 Or separatorPosition >= Len(timingDetail) Then
        Err.Raise vbObjectError + 7622, "VisualTeX", _
            "VisualTeX did not return the equation numbers."
    End If
    VTWordPerformanceNote "requestTiming", Left$(timingDetail, separatorPosition - 1)
    Set result = VTParseDocumentImportManifestText( _
        VTBase64UrlDecodeUtf8(Mid$(timingDetail, separatorPosition + 1)), sessionId)
    VTRequireDispatchValue result, "imageCount"
    If CStr(result("operation")) <> "displayNumbering" Or _
       CLng(result("itemCount")) <> equationCount Or _
       CLng(result("imageCount")) <> imageCount Then
        Err.Raise vbObjectError + 7622, "VisualTeX", _
            "VisualTeX returned equation numbers for another request."
    End If

    If equationCount > 0 Then
        VTRequireDispatchValue result, "docxPath"
        docxPath = CStr(result("docxPath"))
        VTValidateAbsoluteVisualTeXPath docxPath
        Set staging = Documents.Open(FileName:=docxPath, ConfirmConversions:=False, _
            ReadOnly:=True, AddToRecentFiles:=False, Visible:=False)
        stagingDocuments.Add staging
        index = 0
        For Each stagingParagraph In staging.Paragraphs
            index = index + 1
            If index > equationCount Then Exit For
            key = "item" & CStr(index) & "Error"
            If VTCollectionHasKey(result, key) Then
                equationErrors(requestEquations(index)) = _
                    VTBase64UrlDecodeUtf8(CStr(result(key)))
            ElseIf stagingParagraph.Range.OMaths.Count = 1 Then
                Set sources(requestEquations(index)) = VTNativeMathRoot( _
                    stagingParagraph.Range.OMaths(1)).Range.Duplicate
            Else
                Err.Raise vbObjectError + 7622, "VisualTeX", _
                    "Word did not open every numbered equation."
            End If
        Next stagingParagraph
        If index < equationCount Then
            Err.Raise vbObjectError + 7622, "VisualTeX", _
                "Word did not open every numbered equation."
        End If
    End If

    If imageCount > 0 Then
        VTRequireDispatchValue result, "imagesPath"
        imagesPath = CStr(result("imagesPath"))
        Set imageResult = VTReadDocumentImportManifest(imagesPath, sessionId)
        If CStr(imageResult("operation")) <> "displayNumberingImages" Or _
           CLng(imageResult("itemCount")) <> imageCount Then
            Err.Raise vbObjectError + 7622, "VisualTeX", _
                "VisualTeX returned image metadata for another request."
        End If
        For index = 1 To imageCount
            key = "image" & CStr(index) & "Metadata"
            If VTCollectionHasKey(imageResult, key) Then
                If Not VTIsEncodedMetadata(CStr(imageResult(key))) Then
                    Err.Raise vbObjectError + 7622, "VisualTeX", _
                        "VisualTeX returned invalid image formula metadata."
                End If
                imageMetadata(requestImages(index)) = CStr(imageResult(key))
                ' The number takes the formula's letter font.
                key = "image" & CStr(index) & "LetterFont"
                If VTCollectionHasKey(imageResult, key) Then _
                    VTSetWordImageLetterFont doc, _
                        CStr(images(requestImages(index))(1)), CStr(imageResult(key))
            Else
                key = "image" & CStr(index) & "Error"
                VTRequireDispatchValue imageResult, key
                imageErrors(requestImages(index)) = _
                    VTBase64UrlDecodeUtf8(CStr(imageResult(key)))
            End If
        Next index
    End If
End Sub

' --- write ----------------------------------------------------------------------

' Writes every formula's number from the end of the document backwards, so
' every target still to be written keeps its planned address. firstTarget is
' the start of the earliest formula written (no write precedes it).
Private Sub VTWriteDisplayNumbering( _
    ByVal doc As Document, _
    ByVal equations As Collection, _
    ByVal images As Collection, _
    ByRef sources() As Range, _
    ByRef equationErrors() As String, _
    ByRef imageMetadata() As String, _
    ByRef imageErrors() As String, _
    ByRef imagePrefixes() As String, _
    ByRef imageRestarts() As Boolean, _
    ByRef imageValues() As Long, _
    ByRef imageTemplates() As Range, _
    ByRef equationsWritten As Long, _
    ByRef imagesWritten As Long, _
    ByRef failedCount As Long, _
    ByRef firstTarget As Long)

    Dim equationTargets() As Variant
    Dim imageTargets() As Variant
    Dim target As Variant
    Dim equationIndex As Long
    Dim imageIndex As Long
    Dim total As Long
    Dim done As Long
    Dim takeEquation As Boolean

    If equations.Count > 0 Then ReDim equationTargets(1 To equations.Count)
    For Each target In equations
        equationIndex = equationIndex + 1
        equationTargets(equationIndex) = target
    Next target
    If images.Count > 0 Then ReDim imageTargets(1 To images.Count)
    For Each target In images
        imageIndex = imageIndex + 1
        imageTargets(imageIndex) = target
    Next target
    total = equationIndex + imageIndex
    firstTarget = -1
    Do While equationIndex > 0 Or imageIndex > 0
        If equationIndex = 0 Then
            takeEquation = False
        ElseIf imageIndex = 0 Then
            takeEquation = True
        Else
            takeEquation = (equationTargets(equationIndex)(0) > imageTargets(imageIndex)(0))
        End If
        If takeEquation Then
            If sources(equationIndex) Is Nothing Then
                failedCount = failedCount + 1
            Else
                Call VTDisplayNumberingWriteEquation( _
                    doc, equationTargets(equationIndex), sources(equationIndex))
                equationsWritten = equationsWritten + 1
                firstTarget = equationTargets(equationIndex)(0)
            End If
            equationIndex = equationIndex - 1
        Else
            If Len(imageMetadata(imageIndex)) = 0 Then
                failedCount = failedCount + 1
            Else
                VTDisplayNumberingWriteImage doc, imageTargets(imageIndex), _
                    imageMetadata(imageIndex), imagePrefixes(imageIndex), _
                    imageRestarts(imageIndex), imageTemplates(imageIndex)
                imagesWritten = imagesWritten + 1
                firstTarget = imageTargets(imageIndex)(0)
            End If
            imageIndex = imageIndex - 1
        End If
        done = done + 1
        If done Mod 10 = 0 Then
            Application.StatusBar = "VisualTeX: numbering display equations (" & _
                CStr(done) & " of " & CStr(total) & ")..."
        End If
    Loop
End Sub

' Replaces one display equation with its numbered form, as a native edit
' replaces an equation (VTInsertNativeEquationAtRange: one FormattedText
' assignment of the exact OMath, then the equation is found again inside its
' own paragraph), and gives it Windows' display layout at the size it had.
Private Function VTDisplayNumberingWriteEquation( _
    ByVal doc As Document, _
    ByVal target As Variant, _
    ByVal source As Range) As Range

    Dim exact As Range
    Dim paragraphRange As Range
    Dim candidateMath As OMath
    Dim nativeEquation As OMath
    Dim mathIndex As Long
    Dim fontSizePt As Double

    Set exact = doc.Range(target(0), target(1))
    If exact.OMaths.Count < 1 Then GoTo Moved
    Set exact = VTNativeMathRoot(exact.OMaths(1)).Range.Duplicate
    If exact.Start <> target(0) Or exact.End <> target(1) Then GoTo Moved
    ' On Word for Mac the OMath Range of a display equation can take in its
    ' paragraph mark. Replace the equation only: a replaced mark becomes a line
    ' break inside the numbered equation and joins the next paragraph to it.
    ' The paragraph's end tells: Range.Text also reads the end of a structure
    ' that closes the equation (a fraction, a sum) as a CR, and a range that
    ' cuts the structure cannot be replaced (error 6028).
    If exact.End > exact.Start And _
       exact.End = exact.Paragraphs(1).Range.End Then exact.End = exact.End - 1
    VT_WORD_DISPLAY_NUMBERING_WROTE = True
    exact.FormattedText = source.FormattedText
    ' Word for Mac can leave an empty OMath beside the replacement.
    Set paragraphRange = doc.Range(target(0), target(0)).Paragraphs(1).Range.Duplicate
    For mathIndex = paragraphRange.OMaths.Count To 1 Step -1
        Set candidateMath = paragraphRange.OMaths(mathIndex)
        If Not VTOMathHasMeaningfulContent(candidateMath) Then _
            candidateMath.Range.Delete
    Next mathIndex
    If paragraphRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7623, "VisualTeX", _
            "Word did not replace a display equation with one numbered equation."
    End If
    ' Windows' display layout (VTV2FinishDisplayEquation) at the size the
    ' equation had; an equation of mixed sizes keeps them.
    Set nativeEquation = VTNativeMathRoot(paragraphRange.OMaths(1))
    If nativeEquation.Type <> wdOMathDisplay Then nativeEquation.Type = wdOMathDisplay
    Set exact = nativeEquation.Range.Duplicate
    fontSizePt = target(2)
    If fontSizePt > 0# Then VTV2ApplyEquationTypography exact, fontSizePt
    VTV2ConfigureDisplayParagraph exact
    Set nativeEquation = VTNativeMathRoot(exact.OMaths(1))
    Set exact = nativeEquation.Range.Duplicate
    If Not VTEquationHasNumberStructure(nativeEquation, exact) Then
        Err.Raise vbObjectError + 7623, "VisualTeX", _
            "Word did not keep the number of a numbered equation."
    End If
    Set VTDisplayNumberingWriteEquation = exact
    Exit Function

Moved:
    Err.Raise vbObjectError + 7623, "VisualTeX", _
        "A display equation moved before it was numbered. Run batch numbering again."
End Function

' Numbers one image display formula as an edit that changes only its
' numbering does (VTCommitWordDispatch, numberingOnly): the same picture; its
' title, metadata and format record the number; the paragraph number gets
' the prefix and restart planned for the whole batch.
Private Sub VTDisplayNumberingWriteImage( _
    ByVal doc As Document, _
    ByVal target As Variant, _
    ByVal numberedMetadata As String, _
    ByVal prefixText As String, _
    ByVal restartAtOne As Boolean, _
    ByVal template As Range)

    Dim probe As Range
    Dim shape As InlineShape
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim metadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim referenceBaselinePt As Double
    Dim observedWordFontSizePt As Double
    Dim scaleReadable As Boolean
    Dim numberRange As Range

    Set probe = doc.Range(target(0), target(0) + 1)
    If probe.InlineShapes.Count <> 1 Then GoTo Moved
    Set shape = probe.InlineShapes(1)
    If Not VTTryResolveVisualTeXInlineShapeReference( _
       shape, formulaId, displayMode, numbered, metadata, _
       metadataNeedsWrite, formatNeedsWrite, False) Then GoTo Moved
    If StrComp(formulaId, CStr(target(1)), vbTextCompare) <> 0 Then GoTo Moved

    On Error Resume Next
    scaleReadable = VTTryReadWordImageScaleState(doc, formulaId, fontSizePt, _
        referenceWidthPt, referenceHeightPt, referenceBaselinePt, _
        observedWordFontSizePt)
    If Err.Number <> 0 Then scaleReadable = False
    Err.Clear
    On Error GoTo 0

    VT_WORD_DISPLAY_NUMBERING_WROTE = True
    shape.Title = VTFormulaReference(formulaId, "block", True)
    shape.AlternativeText = numberedMetadata
    VTSetWordMetadataPayload doc, formulaId, numberedMetadata
    VTSetWordFormulaFormat doc, formulaId, "block", True
    ' The recount (VTRecountSequenceFields) sets every new number afterwards.
    Set numberRange = VTWriteSingleParagraphImageNumber( _
        shape, formulaId, prefixText, True, restartAtOne, True, _
        template, False)
    VTNormalizeImageDisplayParagraph shape.Range
    If scaleReadable Then
        ' The numbered paragraph style changes the size Word reports for the
        ' picture; record it so the size watcher does not take it for a resize.
        observedWordFontSizePt = VTInlineShapeWordFontSize(shape)
        If Not VTValidWordFormulaFontSize(observedWordFontSizePt) Then _
            observedWordFontSizePt = fontSizePt
        VTSetWordImageScaleState doc, formulaId, fontSizePt, referenceWidthPt, _
            referenceHeightPt, referenceBaselinePt, observedWordFontSizePt
    End If
    VTSetWordImageOwnerBookmark doc, shape, formulaId
    Exit Sub

Moved:
    Err.Raise vbObjectError + 7624, "VisualTeX", _
        "An image formula moved before it was numbered. Run batch numbering again."
End Sub

' --- renumber -------------------------------------------------------------------

' Each family's numbers once, from the earliest new number on (no write
' precedes it, so its position still holds): recounted in one pass
' (VTRecountSequenceFields), or field by field when Word disagrees with the
' count. Then every cross-reference once.
Private Sub VTRenumberDisplayNumbering( _
    ByVal doc As Document, _
    ByVal imagesWritten As Long, _
    ByVal equationsWritten As Long, _
    ByVal transfers As Collection, _
    ByVal restartLosses As Collection, _
    ByRef equationSeries() As String, _
    ByVal headingMarks As Collection, _
    ByVal firstTarget As Long)

    Dim transfer As Range
    Dim candidate As Field
    Dim names() As String
    Dim headings As Collection
    Dim mark As Variant

    If firstTarget < 0 Then Exit Sub
    If imagesWritten > 0 Then
        VTDisplayNumberingDropRestarts doc, restartLosses
        ReDim names(0 To 0)
        names(0) = VT_WORD_IMAGE_SEQUENCE_NAME
        If Not VTRecountSequenceFields(doc, firstTarget, names, Nothing) Then
            VTWordPerformanceNote "imageRecount", "fallback"
            VTUpdateImageNumberFieldsFrom doc, firstTarget, False
        End If
    End If
    If equationsWritten > 0 Then
        For Each transfer In transfers
            VTV2TransferRestart transfer
        Next transfer
        Set headings = New Collection
        If Not headingMarks Is Nothing Then
            For Each mark In headingMarks
                headings.Add Array(mark(0).Start, mark(1))
            Next mark
        End If
        If VTRecountSequenceFields(doc, firstTarget, equationSeries, headings) Then
            VTV2UpdateStyleRefsFrom doc, firstTarget
        Else
            VTWordPerformanceNote "equationRecount", "fallback"
            VTV2UpdateNumbersFrom doc, firstTarget, False
        End If
    End If
    For Each candidate In doc.Fields
        If candidate.Type = wdFieldRef Then candidate.Update
    Next candidate
End Sub

' The SEQ labels of the new equation numbers, once each.
Private Function VTDisplayNumberingSeries(ByRef items() As Variant) As String()
    Dim names() As String
    Dim count As Long
    Dim index As Long
    Dim known As Long
    Dim found As Boolean

    ReDim names(0 To UBound(items) - LBound(items))
    For index = LBound(items) To UBound(items)
        found = False
        For known = 0 To count - 1
            If StrComp(names(known), CStr(items(index)(0)), vbTextCompare) = 0 Then
                found = True
                Exit For
            End If
        Next known
        If Not found Then
            names(count) = CStr(items(index)(0))
            count = count + 1
        End If
    Next index
    ReDim Preserve names(0 To count - 1)
    VTDisplayNumberingSeries = names
End Function

' Every new image number, prepared in a hidden document before Undo starts:
' one number per (chapter prefix, restart) made with its fields
' (VTCreateCanonicalImagePlaceRef evaluates them there, among a handful of
' fields), then one copy of it per shown value with that value written in.
' templates(i) is the number formula i gets copied: no SEQ field is evaluated
' or written in the document for it. Changing another document while the
' document's Undo record is open would split that record.
Private Function VTBuildImageNumberTemplates( _
    ByVal doc As Document, _
    ByRef prefixes() As String, _
    ByRef restarts() As Boolean, _
    ByRef values() As Long, _
    ByRef templateDocument As Document) As Range()

    Dim templates() As Range
    Dim made As New Collection
    Dim bases As New Collection
    Dim level As Long
    Dim separator As String
    Dim prefixText As String
    Dim restartAtOne As Boolean
    Dim baseKey As String
    Dim key As String
    Dim index As Long
    Dim placeRef As Field
    Dim shown As Field
    Dim base As Variant
    Dim copiedRange As Range

    ReDim templates(LBound(prefixes) To UBound(prefixes))
    level = VTEquationNumberingRestartLevel(doc)
    If level > 0 Then separator = VTEquationNumberingDisplaySeparator(doc)
    Set templateDocument = Documents.Add(Visible:=False)
    For index = LBound(prefixes) To UBound(prefixes)
        If level > 0 Then
            prefixText = prefixes(index) & separator
            restartAtOne = restarts(index)
        End If
        baseKey = VTImageNumberTemplateKey(prefixText, restartAtOne)
        key = baseKey & "|" & CStr(values(index))
        If Not VTCollectionHasKey(made, key) Then
            If Not VTCollectionHasKey(bases, baseKey) Then
                Set placeRef = VTCreateCanonicalImagePlaceRef( _
                    templateDocument, templateDocument.Content.End - 1, _
                    restartAtOne, prefixText)
                bases.Add Array(templateDocument.Range( _
                    VTEquationFieldStart(placeRef), VTEquationFieldEnd(placeRef))), baseKey
                templateDocument.Content.InsertParagraphAfter
            End If
            base = bases(baseKey)
            Set copiedRange = templateDocument.Range( _
                templateDocument.Content.End - 1, templateDocument.Content.End - 1)
            copiedRange.FormattedText = base(0).FormattedText
            Set placeRef = VTImagePlaceRef(templateDocument.Paragraphs.Last.Range)
            Set shown = VTImageSequenceField(placeRef, False)
            If placeRef Is Nothing Or shown Is Nothing Then
                Err.Raise vbObjectError + 7625, "VisualTeX", _
                    "Word could not prepare an image Equation number."
            End If
            If StrComp(shown.Result.Text, CStr(values(index)), vbBinaryCompare) <> 0 Then _
                shown.Result.Text = CStr(values(index))
            made.Add Array(templateDocument.Range( _
                VTEquationFieldStart(placeRef), VTEquationFieldEnd(placeRef))), key
            templateDocument.Content.InsertParagraphAfter
        End If
        Set templates(index) = made(key)(0)
    Next index
    doc.Activate
    VTBuildImageNumberTemplates = templates
End Function

Private Function VTImageNumberTemplateKey( _
    ByVal prefixText As String, ByVal restartAtOne As Boolean) As String
    VTImageNumberTemplateKey = IIf(restartAtOne, "r|", "n|") & prefixText
End Function

' Removes \r 1 from the hidden SEQ field of each image in losses
' (VTPlanDisplayNumberingImageNumbers), as VTRepairImageNumberRestartsFrom
' rewrites a restart that no longer holds.
Private Sub VTDisplayNumberingDropRestarts( _
    ByVal doc As Document, ByVal losses As Collection)

    Dim loss As Range
    Dim placeRef As Field
    Dim hiddenField As Field

    For Each loss In losses
        If loss.InlineShapes.Count = 1 Then
            Set placeRef = VTImagePlaceRef(loss.InlineShapes(1).Range)
            If Not placeRef Is Nothing Then
                Set hiddenField = VTImageSequenceField(placeRef, True)
                If Not hiddenField Is Nothing Then
                    If VTEquationSequenceFieldHasSwitch(hiddenField, "r") Then
                        hiddenField.Code.Text = " " & VTImageSequenceInstruction( _
                            doc, True, False, False) & " "
                    End If
                End If
            End If
        End If
    Next loss
End Sub

' --- recount ---------------------------------------------------------------------

' One family's equation numbers after a batch write, set in one pass over the
' document's SEQ fields (Windows' computed results): Word evaluates a SEQ
' field by counting every field before it, so updating n numbers field by
' field takes O(n^2) time. Every field of the series in seriesNames is counted
' as Word counts it (\r N restarts, \s L restarts after a heading of level L
' or above, \h increments without showing, \c shows the current value), and
' each shown result from fromPosition on that differs is written directly.
' Word itself then updates a sample (the first field of each restart scope,
' spread over at most VT_WORD_DISPLAY_NUMBERING_PROBES, and the last field).
' True only when Word agrees with every sampled field; the caller then has
' nothing left to update, and updates every field otherwise.
Private Function VTRecountSequenceFields( _
    ByVal doc As Document, _
    ByVal fromPosition As Long, _
    ByRef seriesNames() As String, _
    ByVal headings As Collection) As Boolean

    Dim headingStarts() As Long
    Dim headingLevels() As Long
    Dim headingCount As Long
    Dim headingIndex As Long
    Dim counters() As Long
    Dim lowest() As Long
    Dim entry As Variant
    Dim visited As New Collection
    Dim scopeFields As New Collection
    Dim scopeValues As New Collection
    Dim candidate As Field
    Dim nested As Field
    Dim lastField As Field
    Dim lastValue As Long
    Dim index As Long
    Dim stride As Long

    If Not headings Is Nothing Then
        If headings.Count > 0 Then
            ReDim headingStarts(1 To headings.Count)
            ReDim headingLevels(1 To headings.Count)
            For Each entry In headings
                headingCount = headingCount + 1
                headingStarts(headingCount) = entry(0)
                headingLevels(headingCount) = entry(1)
            Next entry
        End If
    End If
    ReDim counters(LBound(seriesNames) To UBound(seriesNames))
    ReDim lowest(LBound(seriesNames) To UBound(seriesNames))
    For index = LBound(seriesNames) To UBound(seriesNames)
        lowest(index) = 99
    Next index
    headingIndex = 1

    ' Nested fields (an image number's SEQ fields inside its MACROBUTTON) are
    ' counted in document order through their parent, once each.
    For Each candidate In doc.Content.Fields
        If VTIsImagePlaceRef(candidate) Then
            For Each nested In candidate.Code.Fields
                If Not VTRecountSequenceField(nested, fromPosition, seriesNames, _
                   counters, lowest, headingStarts, headingLevels, headingCount, _
                   headingIndex, visited, scopeFields, scopeValues, _
                   lastField, lastValue) Then Exit Function
            Next nested
        ElseIf candidate.Type = wdFieldSequence Then
            If Not VTRecountSequenceField(candidate, fromPosition, seriesNames, _
               counters, lowest, headingStarts, headingLevels, headingCount, _
               headingIndex, visited, scopeFields, scopeValues, _
               lastField, lastValue) Then Exit Function
        End If
    Next candidate

    ' Word's own evaluation of the sample.
    If scopeFields.Count > 0 Then
        stride = (scopeFields.Count + VT_WORD_DISPLAY_NUMBERING_PROBES - 1) \ _
            VT_WORD_DISPLAY_NUMBERING_PROBES
        For index = 1 To scopeFields.Count Step stride
            If Not VTSequenceFieldShows(scopeFields(index), CLng(scopeValues(index))) Then _
                Exit Function
        Next index
    End If
    If Not lastField Is Nothing Then
        If Not VTSequenceFieldShows(lastField, lastValue) Then Exit Function
    End If
    VTRecountSequenceFields = True
End Function

' Word's own evaluation of one SEQ field equals value.
Private Function VTSequenceFieldShows( _
    ByVal candidate As Field, ByVal value As Long) As Boolean

    candidate.Update
    VTSequenceFieldShows = (StrComp(Trim$(candidate.Result.Text), _
        CStr(value), vbBinaryCompare) = 0)
End Function

' One SEQ field of VTRecountSequenceFields: counts it and, from fromPosition
' on, writes its shown result when it differs. False when the field cannot be
' counted here (an unsupported switch or number format).
Private Function VTRecountSequenceField( _
    ByVal candidate As Field, _
    ByVal fromPosition As Long, _
    ByRef seriesNames() As String, _
    ByRef counters() As Long, _
    ByRef lowest() As Long, _
    ByRef headingStarts() As Long, _
    ByRef headingLevels() As Long, _
    ByVal headingCount As Long, _
    ByRef headingIndex As Long, _
    ByVal visited As Collection, _
    ByVal scopeFields As Collection, _
    ByVal scopeValues As Collection, _
    ByRef lastField As Field, _
    ByRef lastValue As Long) As Boolean

    Dim instruction As String
    Dim fieldStart As Long
    Dim series As Long
    Dim index As Long
    Dim level As Long
    Dim levelPosition As Long
    Dim value As Long
    Dim hidden As Boolean
    Dim current As Boolean

    VTRecountSequenceField = True
    If candidate.Type <> wdFieldSequence Then Exit Function
    fieldStart = VTEquationFieldStart(candidate)
    If VTCollectionHasKey(visited, CStr(fieldStart)) Then Exit Function
    visited.Add True, CStr(fieldStart)
    instruction = candidate.Code.Text
    series = LBound(seriesNames) - 1
    For index = LBound(seriesNames) To UBound(seriesNames)
        If StrComp(VTV2SequenceName(instruction), seriesNames(index), vbTextCompare) = 0 Then
            series = index
            Exit For
        End If
    Next index
    ' A heading passed since a series' previous field restarts its \s numbers.
    Do While headingIndex <= headingCount
        If headingStarts(headingIndex) >= fieldStart Then Exit Do
        For index = LBound(lowest) To UBound(lowest)
            If headingLevels(headingIndex) < lowest(index) Then _
                lowest(index) = headingLevels(headingIndex)
        Next index
        headingIndex = headingIndex + 1
    Loop
    If series < LBound(seriesNames) Then Exit Function

    VTRecountSequenceField = False
    If InStr(1, instruction, "\*", vbBinaryCompare) > 0 And _
       InStr(1, instruction, "\* ARABIC", vbTextCompare) = 0 Then Exit Function
    If InStr(1, instruction, "\n", vbTextCompare) > 0 Then Exit Function
    hidden = VTEquationSequenceFieldHasSwitch(candidate, "h")
    current = VTEquationSequenceFieldHasSwitch(candidate, "c")
    levelPosition = InStr(1, instruction, "\s ", vbTextCompare)
    If levelPosition > 0 Then level = CLng(Val(Mid$(instruction, levelPosition + 3, 2)))
    If current Then
        value = counters(series)
    ElseIf VTV2RestartValue(instruction) > 0 Then
        value = VTV2RestartValue(instruction)
    ElseIf level > 0 And lowest(series) <= level Then
        value = 1
    Else
        value = counters(series) + 1
    End If
    If Not current Then
        counters(series) = value
        lowest(series) = 99
    End If
    If fieldStart >= fromPosition And Not hidden Then
        If StrComp(Trim$(candidate.Result.Text), CStr(value), vbBinaryCompare) <> 0 Then _
            candidate.Result.Text = CStr(value)
        If value = 1 Or scopeFields.Count = 0 Then
            scopeFields.Add candidate
            scopeValues.Add value
        End If
        Set lastField = candidate
        lastValue = value
    End If
    VTRecountSequenceField = True
End Function
