Attribute VB_Name = "VTPowerPointNativeMath"
Option Explicit

Private NativeTargets As Collection
Private Const NativeHost As String = "powerpoint"
Private Const TextPrefix As String = "visualtex-ppt-native-text:"
Private Const SlidePrefix As String = "visualtex-ppt-native-slide:"
Private NativeApplyBusy As Boolean
' The table cell VTTryResolveNativeMathZone found the equation in (Nothing: a text box).
Private ResolvedTableShape As PowerPoint.Shape
Private ResolvedCellRow As Long
Private ResolvedCellColumn As Long

Public Sub VisualTeX_NewInlineNativeEquation()
    VTNewPowerPointNativeEquation "inline"
End Sub

Public Sub VisualTeX_NewDisplayNativeEquation()
    VTNewPowerPointNativeEquation "block"
End Sub

Private Function VTNativeTarget(ByVal sessionId As String) As VTPowerPointNativeTarget
    If NativeTargets Is Nothing Then GoTo MissingTarget
    On Error GoTo MissingTarget
    Set VTNativeTarget = NativeTargets(sessionId)
    Exit Function
MissingTarget:
    Err.Raise vbObjectError + 7801, "VisualTeX", _
        "The original PowerPoint insertion position is no longer available. Start a new insertion at the intended text position."
End Function

Public Function VTNativeCaptureTarget(ByVal sessionId As String) As VTPowerPointNativeTarget
    Dim target As New VTPowerPointNativeTarget
    Dim selected As PowerPoint.Selection
    Dim selectedText As PowerPoint.TextRange
    Dim fullText As PowerPoint.TextRange
    Dim candidate As PowerPoint.Shape
    Dim fontSize As Double
    Dim index As Long
    Dim selectedEnd As Long
    Dim cellRow As Long
    Dim cellColumn As Long

    VTRequireWritablePowerPointPresentation
    Set target.Owner = ActivePresentation
    Set target.OwnerWindow = ActiveWindow
    Set target.TargetSlide = ActiveWindow.View.Slide
    target.SlideId = target.TargetSlide.SlideID
    target.SessionId = sessionId
    Set selected = ActiveWindow.Selection
    If selected.Type = ppSelectionText Then
        If selected.ShapeRange.Count <> 1 Then
            Err.Raise vbObjectError + 7802, "VisualTeX", "Place the text cursor in exactly one text box."
        End If
        Set candidate = selected.ShapeRange(1)
        If candidate.HasTextFrame <> msoTrue And candidate.HasTable = msoTrue Then
            ' The caret in a table cell: the equation goes into that cell's text.
            If Not VTSelectedTableCell(candidate, cellRow, cellColumn) Then
                Err.Raise vbObjectError + 7802, "VisualTeX", "Place the text cursor in exactly one table cell."
            End If
            Set target.TableShape = candidate
            target.CellRow = cellRow
            target.CellColumn = cellColumn
            Set target.TargetShape = candidate.Table.Cell(cellRow, cellColumn).Shape
        Else
            Set target.TargetShape = candidate
        End If
        Set selectedText = selected.TextRange
    ElseIf selected.Type = ppSelectionShapes Then
        If selected.ShapeRange.Count = 1 Then
            Set candidate = selected.ShapeRange(1)
            If candidate.HasTextFrame = msoTrue Then
                If candidate.TextFrame.TextRange.Length = 0 Then
                    Set target.TargetShape = candidate
                    Set selectedText = candidate.TextFrame.TextRange.Characters(1, 0)
                End If
            End If
        End If
    End If
    If Not target.TargetShape Is Nothing Then
        If target.TargetShape.HasTextFrame <> msoTrue Then
            Err.Raise vbObjectError + 7802, "VisualTeX", "The current selection is not an editable text box."
        End If
        Set fullText = target.TargetShape.TextFrame.TextRange
        target.ShapeId = VTNativeHostShape(target).Id
        target.TextStart = selectedText.Start
        target.TextLength = selectedText.Length
        target.OriginalText = fullText.Text
        If target.TextStart < 1 Or target.TextLength < 0 Or _
           target.TextStart + target.TextLength > fullText.Length + 1 Then
            Err.Raise vbObjectError + 7802, "VisualTeX", "PowerPoint returned an invalid text insertion position."
        End If
        selectedEnd = target.TextStart + target.TextLength
        ' A cell's equation is written as text (VTWriteCellEquation adds a display
        ' equation's paragraph breaks itself): no paragraph context is captured.
        If target.TableShape Is Nothing Then
            If target.TextStart > 1 Then target.PrefixBreak = (Mid$(target.OriginalText, target.TextStart - 1, 1) <> vbCr)
            If selectedEnd <= Len(target.OriginalText) Then target.SuffixBreak = (Mid$(target.OriginalText, selectedEnd, 1) <> vbCr)
        End If
        On Error Resume Next
        fontSize = selectedText.Font.Size
        If fontSize <= 0# Then fontSize = selectedText.Characters(1, 1).Font.Size
        If fontSize <= 0# Then fontSize = fullText.Font.Size
        On Error GoTo 0
    End If
    If fontSize < 1# Or fontSize > 512# Then fontSize = 18#
    target.FontSizePt = fontSize
    If NativeTargets Is Nothing Then Set NativeTargets = New Collection
    For index = NativeTargets.Count To 1 Step -1
        If NativeTargets(index).Completed Then NativeTargets.Remove index
    Next index
    NativeTargets.Add target, sessionId
    Set VTNativeCaptureTarget = target
End Function

' MathZones() without indices describes the collection, not one formula range.
' Always resolve MathZones(index, 1); its Start/Length follow Office's current
' native equation boundaries, including equations Office naturally merged.
Public Function VTTryResolveNativeMathZone( _
    ByVal selected As PowerPoint.Selection, _
    ByRef targetShape As PowerPoint.Shape, ByRef mathZone As Object, _
    Optional ByVal useScreenPoint As Boolean = False, _
    Optional ByVal screenX As Double = 0#, Optional ByVal screenY As Double = 0#) As Boolean

    Dim fullText As Object
    Dim candidate As Object
    Dim selectionStart As Long
    Dim selectionEnd As Long
    Dim index As Long
    Dim matches As Long
    Dim hit As Boolean
    Dim windowObject As Object
    Dim left As Double, top As Double, right As Double, bottom As Double
    Dim offsetX As Double, offsetY As Double
    Dim cellRow As Long, cellColumn As Long
    On Error GoTo NotMath
    Set targetShape = Nothing
    Set mathZone = Nothing
    Set ResolvedTableShape = Nothing
    If selected Is Nothing Then Exit Function
    If selected.Type <> ppSelectionText And selected.Type <> ppSelectionShapes Then Exit Function
    If selected.ShapeRange.Count <> 1 Then Exit Function
    Set targetShape = selected.ShapeRange(1)
    If targetShape.HasTextFrame <> msoTrue Then
        ' The caret in a table cell: that cell's equations.
        If selected.Type <> ppSelectionText Or targetShape.HasTable <> msoTrue Then GoTo NotMath
        If Not VTSelectedTableCell(targetShape, cellRow, cellColumn) Then GoTo NotMath
        Set ResolvedTableShape = targetShape
        ResolvedCellRow = cellRow
        ResolvedCellColumn = cellColumn
        ' A cell's text reports its bounds from the table's corner.
        offsetX = targetShape.Left
        offsetY = targetShape.Top
        Set targetShape = targetShape.Table.Cell(cellRow, cellColumn).Shape
    End If
    Set fullText = targetShape.TextFrame2.TextRange
    If selected.Type = ppSelectionText Then
        selectionStart = selected.TextRange.Start
        selectionEnd = selectionStart + selected.TextRange.Length
    End If
    Set windowObject = ActiveWindow
    For index = 1 To fullText.MathZones.Count
        Set candidate = fullText.MathZones(index, 1)
        hit = False
        If useScreenPoint Then
            left = windowObject.PointsToScreenPixelsX(offsetX + candidate.BoundLeft)
            top = windowObject.PointsToScreenPixelsY(offsetY + candidate.BoundTop)
            right = windowObject.PointsToScreenPixelsX(offsetX + candidate.BoundLeft + candidate.BoundWidth)
            bottom = windowObject.PointsToScreenPixelsY(offsetY + candidate.BoundTop + candidate.BoundHeight)
            hit = (screenX >= left - 1# And screenX <= right + 1# And _
                   screenY >= top - 1# And screenY <= bottom + 1#)
        ElseIf selected.Type = ppSelectionText Then
            If selectionStart = selectionEnd Then
                hit = (selectionStart >= candidate.Start And selectionStart < candidate.Start + candidate.Length)
            Else
                hit = (selectionStart < candidate.Start + candidate.Length And selectionEnd > candidate.Start)
            End If
        Else
            hit = True
        End If
        If hit And candidate.Length > 0 Then
            matches = matches + 1
            Set mathZone = candidate
        End If
    Next index
    If matches <> 1 Then GoTo NotMath
    VTTryResolveNativeMathZone = True
    Exit Function
NotMath:
    Set targetShape = Nothing
    Set mathZone = Nothing
    Set ResolvedTableShape = Nothing
End Function

' The table cell that holds the text caret or text selection: PowerPoint marks
' exactly that cell Selected (every grid position of a merged cell, all with the
' merged cell's frame; the first is the one whose text it is). False when no cell
' or several cells are selected.
Private Function VTSelectedTableCell(ByVal tableShape As PowerPoint.Shape, _
    ByRef cellRow As Long, ByRef cellColumn As Long) As Boolean
    Dim r As Long, c As Long
    Dim cellShape As PowerPoint.Shape
    Dim found As Boolean
    Dim frameLeft As Single, frameTop As Single, frameWidth As Single, frameHeight As Single
    On Error GoTo NoCell
    cellRow = 0
    cellColumn = 0
    If tableShape.HasTable <> msoTrue Then Exit Function
    For r = 1 To tableShape.Table.Rows.Count
        For c = 1 To tableShape.Table.Columns.Count
            If tableShape.Table.Cell(r, c).Selected Then
                Set cellShape = tableShape.Table.Cell(r, c).Shape
                If Not found Then
                    found = True
                    cellRow = r
                    cellColumn = c
                    frameLeft = cellShape.Left: frameTop = cellShape.Top
                    frameWidth = cellShape.Width: frameHeight = cellShape.Height
                ElseIf Abs(cellShape.Left - frameLeft) > 0.5 Or Abs(cellShape.Top - frameTop) > 0.5 Or _
                       Abs(cellShape.Width - frameWidth) > 0.5 Or Abs(cellShape.Height - frameHeight) > 0.5 Then
                    GoTo NoCell
                End If
            End If
        Next c
    Next r
    VTSelectedTableCell = found
    Exit Function
NoCell:
    cellRow = 0
    cellColumn = 0
    VTSelectedTableCell = False
End Function

' The slide shape holding the target's text: the text box, or the table of a cell.
Private Function VTNativeHostShape(ByVal target As VTPowerPointNativeTarget) As PowerPoint.Shape
    If target.TableShape Is Nothing Then
        Set VTNativeHostShape = target.TargetShape
    Else
        Set VTNativeHostShape = target.TableShape
    End If
End Function

Private Function VTNativeCellJson(ByVal target As VTPowerPointNativeTarget) As String
    If Not target.TableShape Is Nothing Then VTNativeCellJson = ",""tableCell"":true"
End Function

Private Function VTIsOrdinaryNativeBoundarySpace(ByVal fullText As Object, ByVal index As Long) As Boolean
    Dim zone As Object
    Dim zoneIndex As Long
    Dim character As String
    If index < 1 Or index > Len(fullText.Text) Then Exit Function
    character = Mid$(fullText.Text, index, 1)
    If character <> " " And character <> vbTab And character <> ChrW(160) Then Exit Function
    For zoneIndex = 1 To fullText.MathZones.Count
        Set zone = fullText.MathZones(zoneIndex, 1)
        If index >= zone.Start And index < zone.Start + zone.Length Then Exit Function
    Next zoneIndex
    VTIsOrdinaryNativeBoundarySpace = True
End Function

Public Function VTNativeCaptureEditTarget( _
    ByVal sessionId As String, ByVal selectedShape As PowerPoint.Shape, _
    ByVal mathZone As Object) As VTPowerPointNativeTarget
    Dim target As New VTPowerPointNativeTarget
    Dim index As Long
    VTRequireWritablePowerPointPresentation
    Set target.Owner = ActivePresentation
    Set target.OwnerWindow = ActiveWindow
    Set target.TargetSlide = ActiveWindow.View.Slide
    Set target.TargetShape = selectedShape
    target.SessionId = sessionId
    target.SlideId = target.TargetSlide.SlideID
    If Not ResolvedTableShape Is Nothing Then
        ' The equation is in the table cell VTTryResolveNativeMathZone found.
        Set target.TableShape = ResolvedTableShape
        target.CellRow = ResolvedCellRow
        target.CellColumn = ResolvedCellColumn
    End If
    target.ShapeId = VTNativeHostShape(target).Id
    target.TextStart = mathZone.Start
    target.TextLength = mathZone.Length
    target.OriginalText = selectedShape.TextFrame.TextRange.Text
    target.IsEdit = True
    ' PowerPoint's smart replacement of a math-only range consumes adjacent
    ' ordinary spaces. Replace the exact visible whitespace together with the
    ' equation using Office's own native text fragment, preserving its styles.
    ' Never include another MathZone, paragraph mark, or non-whitespace text.
    ' A cell's equation is replaced as text: exactly its own range.
    Dim fullText As Object
    Dim rangeEnd As Long
    Set fullText = selectedShape.TextFrame2.TextRange
    target.ReplaceStart = target.TextStart
    rangeEnd = target.TextStart + target.TextLength
    If target.TableShape Is Nothing Then
        Do While VTIsOrdinaryNativeBoundarySpace(fullText, target.ReplaceStart - 1)
            target.ReplaceStart = target.ReplaceStart - 1
        Loop
        Do While VTIsOrdinaryNativeBoundarySpace(fullText, rangeEnd)
            rangeEnd = rangeEnd + 1
        Loop
    End If
    target.ReplaceLength = rangeEnd - target.ReplaceStart
    target.PrefixBreak = False
    target.SuffixBreak = False
    target.FontSizePt = mathZone.Font.Size
    If target.FontSizePt < 1# Or target.FontSizePt > 512# Then target.FontSizePt = mathZone.Characters(1, 1).Font.Size
    If target.FontSizePt < 1# Or target.FontSizePt > 512# Then target.FontSizePt = 18#
    If NativeTargets Is Nothing Then Set NativeTargets = New Collection
    For index = NativeTargets.Count To 1 Step -1
        If NativeTargets(index).Completed Then NativeTargets.Remove index
    Next index
    NativeTargets.Add target, sessionId
    Set VTNativeCaptureEditTarget = target
End Function

Public Function VTNativeEditRequest(ByVal target As VTPowerPointNativeTarget) As String
    Dim geometry As String
    Dim sourceId As String
    Dim layout As String
    Dim paragraph As Object
    Dim fullText As Object
    Dim index As Long
    Dim beforeLength As Long, afterLength As Long
    Dim outsideText As String
    layout = "inline"
    Set fullText = target.TargetShape.TextFrame2.TextRange
    For index = 1 To fullText.Paragraphs.Count
        Set paragraph = fullText.Paragraphs(index, 1)
        If paragraph.Start <= target.TextStart And paragraph.Start + paragraph.Length >= target.TextStart + target.TextLength Then
            beforeLength = target.TextStart - paragraph.Start
            afterLength = paragraph.Start + paragraph.Length - target.TextStart - target.TextLength
            If beforeLength > 0 Then outsideText = fullText.Characters(paragraph.Start, beforeLength).Text
            If afterLength > 0 Then outsideText = outsideText & fullText.Characters(target.TextStart + target.TextLength, afterLength).Text
            If Len(Trim$(Replace$(outsideText, vbCr, ""))) = 0 Then layout = "block"
            Exit For
        End If
    Next index
    geometry = VTPowerPointGeometryJson(target.TargetSlide, VTNativeHostShape(target), target.FontSizePt)
    geometry = Left$(geometry, Len(geometry) - 1) & ",""nativeTextTarget"":{" & _
        """rangeStart"":" & CStr(target.TextStart) & ",""rangeLength"":" & CStr(target.TextLength) & "," & _
        """leadingParagraph"":false,""trailingParagraph"":false" & VTNativeCellJson(target) & "}}"
    sourceId = TextPrefix & CStr(target.SlideId) & ":" & CStr(target.ShapeId) & ":" & _
        CStr(target.TextStart) & ":" & CStr(target.TextLength)
    VTNativeEditRequest = VTRequestJson(target.SessionId, NativeHost, "edit", "", layout, False, _
        VTPresentationIdentityFor(target.Owner), sourceId, "", "", geometry, True)
End Function

Public Function VTEditSelectedNativeEquation( _
    Optional ByVal useScreenPoint As Boolean = False, _
    Optional ByVal screenX As Double = 0#, Optional ByVal screenY As Double = 0#) As Boolean
    Dim selectedShape As Shape
    Dim mathZone As Object
    Dim target As VTPowerPointNativeTarget
    Dim sessionId As String
    Dim request As String
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    If NativeApplyBusy Then Exit Function
    If ActiveWindow Is Nothing Then Exit Function
    If Not VTTryResolveNativeMathZone(ActiveWindow.Selection, selectedShape, mathZone, useScreenPoint, screenX, screenY) Then Exit Function
    VTEditSelectedNativeEquation = True
    On Error GoTo Failed
    sessionId = VTNewUuidV4()
    Set target = VTNativeCaptureEditTarget(sessionId, selectedShape, mathZone)
    request = VTNativeEditRequest(target)
    Call VTWriteAndLaunchSession(NativeHost, sessionId, request)
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If Not NativeTargets Is Nothing Then NativeTargets.Remove sessionId
    If Len(sessionId) > 0 Then VTDeleteSessionFiles sessionId
    On Error GoTo 0
    If Not useScreenPoint Then VTShowError "PowerPoint native equation editing", errorNumber, errorText
End Function

Public Sub VisualTeX_EditNativeAtScreenPoint(ByVal xText As String, ByVal yText As String)
    Dim handled As Boolean
    On Error GoTo Finished
    handled = VTEditSelectedNativeEquation(True, VTParseInvariantDouble(xText), VTParseInvariantDouble(yText))
Finished:
End Sub

Public Sub VisualTeX_CopyPowerPointNativeEditSource()
    Dim sessionId As String
    Dim target As VTPowerPointNativeTarget
    Dim fullText As Object
    Dim zone As Object
    Dim index As Long
    Dim resultPath As String
    Dim errorText As String
    sessionId = VTReadActiveSessionId(NativeHost)
    resultPath = VTSessionDirectory(sessionId) & "/native-edit-copy-status.txt"
    On Error GoTo Failed
    Set target = VTNativeTarget(sessionId)
    If target.Completed Then
        VTWriteTextAtomic resultPath, "completed"
        Exit Sub
    End If
    If Not target.IsEdit Then Err.Raise 5, "VisualTeX", "The session has no original native equation."
    VTNativeValidateTarget target
    Set fullText = target.TargetShape.TextFrame2.TextRange
    For index = 1 To fullText.MathZones.Count
        Set zone = fullText.MathZones(index, 1)
        If zone.Start = target.TextStart And zone.Length = target.TextLength Then
            fullText.Characters(target.ReplaceStart, target.ReplaceLength).Copy
            VTWriteTextAtomic resultPath, "ok"
            Exit Sub
        End If
    Next index
    Err.Raise vbObjectError + 7803, "VisualTeX", "Office changed the native equation boundary while the editor was open. Reopen the current equation."
Failed:
    errorText = Err.Description
    On Error Resume Next
    VTWriteTextAtomic resultPath, "error" & vbLf & errorText
    On Error GoTo 0
End Sub

Private Sub VTNativeValidateTarget(ByVal target As VTPowerPointNativeTarget)
    Dim owner As PowerPoint.Presentation
    Dim currentSlide As PowerPoint.Slide
    Dim shapeItem As PowerPoint.Shape
    Dim ownerFound As Boolean
    Dim shapeFound As Boolean

    If target.Completed Then Exit Sub
    For Each owner In Presentations
        If owner Is target.Owner Then ownerFound = True: Exit For
    Next owner
    If Not ownerFound Then GoTo StaleTarget
    VTRequireWritablePowerPointPresentationObject target.Owner
    Set currentSlide = target.Owner.Slides.FindBySlideID(target.SlideId)
    If Not currentSlide Is target.TargetSlide Then GoTo StaleTarget
    If Not target.TargetShape Is Nothing Then
        For Each shapeItem In currentSlide.Shapes
            If shapeItem Is VTNativeHostShape(target) Then
                If shapeItem.Id = target.ShapeId Then shapeFound = True
                Exit For
            End If
        Next shapeItem
        If Not shapeFound Then GoTo StaleTarget
        If Not target.TableShape Is Nothing Then
            If target.TableShape.HasTable <> msoTrue Then GoTo StaleTarget
            On Error GoTo StaleTarget
            Set target.TargetShape = target.TableShape.Table.Cell(target.CellRow, target.CellColumn).Shape
            On Error GoTo 0
        End If
        If target.TargetShape.HasTextFrame <> msoTrue Then GoTo StaleTarget
        If StrComp(target.TargetShape.TextFrame.TextRange.Text, target.OriginalText, vbBinaryCompare) <> 0 Then
            Err.Raise vbObjectError + 7803, "VisualTeX", _
                "The text box changed while the formula editor was open. No equation was inserted; start again at the intended text position."
        End If
    End If
    Exit Sub
StaleTarget:
    Err.Raise vbObjectError + 7803, "VisualTeX", _
        "The original PowerPoint presentation, slide, or text box is no longer available. No equation was inserted."
End Sub

Private Sub VTNewPowerPointNativeEquation(ByVal displayMode As String)
    Dim sessionId As String
    Dim target As VTPowerPointNativeTarget
    Dim request As String
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo Failed
    sessionId = VTNewUuidV4()
    Set target = VTNativeCaptureTarget(sessionId)
    request = VTNativeCreateRequest(target, displayMode)
    Call VTWriteAndLaunchSession(NativeHost, sessionId, request)
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If Not NativeTargets Is Nothing Then NativeTargets.Remove sessionId
    If Len(sessionId) > 0 Then VTDeleteSessionFiles sessionId
    On Error GoTo 0
    VTShowError "PowerPoint native equation insertion", errorNumber, errorText
End Sub

Public Function VTNativeCreateRequest(ByVal target As VTPowerPointNativeTarget, ByVal displayMode As String) As String
    Dim geometry As String
    Dim sourceId As String
    If Not target.TargetShape Is Nothing Then
        geometry = VTPowerPointGeometryJson(target.TargetSlide, VTNativeHostShape(target), target.FontSizePt)
        geometry = Left$(geometry, Len(geometry) - 1) & ",""nativeTextTarget"":{" & _
            """rangeStart"":" & CStr(target.TextStart) & ",""rangeLength"":" & CStr(target.TextLength) & "," & _
            """leadingParagraph"":" & IIf(target.PrefixBreak, "true", "false") & "," & _
            """trailingParagraph"":" & IIf(target.SuffixBreak, "true", "false") & VTNativeCellJson(target) & "}}"
        sourceId = TextPrefix & CStr(target.SlideId) & ":" & CStr(target.ShapeId) & ":" & _
            CStr(target.TextStart) & ":" & CStr(target.TextLength)
    Else
        geometry = "{" & _
            """presentationIdentity"":" & VTJsonString(VTPresentationIdentityFor(target.Owner)) & "," & _
            """slideIndex"":" & CStr(target.TargetSlide.SlideIndex) & "," & _
            """slideId"":" & CStr(target.SlideId) & "," & _
            """shapeIndex"":0,""shapeId"":0,""shapeName"":""""," & _
            """left"":0,""top"":0,""width"":240,""height"":90," & _
            """rotation"":0,""zOrder"":0,""fontSizePt"":" & VTJsonNumber(target.FontSizePt) & "}"
        sourceId = SlidePrefix & CStr(target.SlideId)
    End If
    VTNativeCreateRequest = VTRequestJson(target.SessionId, NativeHost, "create", "", displayMode, False, _
        VTPresentationIdentityFor(target.Owner), sourceId, "", "", geometry, True)
End Function

Public Sub VisualTeX_CopyNativeParagraphContext()
    Dim sessionId As String
    Dim target As VTPowerPointNativeTarget
    Dim fullText As PowerPoint.TextRange
    Dim paragraph As PowerPoint.TextRange
    Dim index As Long
    Dim resultPath As String
    Dim errorText As String
    sessionId = VTReadActiveSessionId(NativeHost)
    resultPath = VTSessionDirectory(sessionId) & "/native-context-status.txt"
    On Error GoTo Failed
    Set target = VTNativeTarget(sessionId)
    If target.Completed Then
        VTWriteTextAtomic resultPath, "completed"
        Exit Sub
    End If
    VTNativeValidateTarget target
    If target.TargetShape Is Nothing Or Not target.PrefixBreak Then
        Err.Raise 5, "VisualTeX", "No leading text paragraph requires preservation."
    End If
    Set fullText = target.TargetShape.TextFrame.TextRange
    For index = 1 To fullText.Paragraphs.Count
        Set paragraph = fullText.Paragraphs(index, 1)
        If paragraph.Start <= target.TextStart And _
           target.TextStart <= paragraph.Start + paragraph.Length Then
            ' TextRange.Copy drops paragraph properties at an implicit final
            ' paragraph mark. Shape.Copy retains the complete native text style;
            ' the backend extracts only this paragraph's formatting, never its
            ' shape geometry, metadata or original text.
            target.TargetShape.Copy
            VTWriteTextAtomic VTSessionDirectory(sessionId) & "/native-context-paragraph.txt", CStr(index)
            VTWriteTextAtomic resultPath, "ok"
            Exit Sub
        End If
    Next index
    Err.Raise 5, "VisualTeX", "The original text paragraph is no longer available."
Failed:
    errorText = Err.Description
    On Error Resume Next
    VTWriteTextAtomic resultPath, "error" & vbLf & errorText
    On Error GoTo 0
End Sub

' Return errors through the existing local session channel instead of leaving an
' unhandled VBA dialog blocking the desktop app's synchronous Apple event.
Public Sub VTDispatchPowerPointNativeEquation(ByVal sessionId As String, ByVal dispatch As Object)
    Dim resultPath As String
    Dim errorText As String
    Dim errorNumber As Long
    resultPath = VTSessionDirectory(sessionId) & "/native-result.txt"
    On Error GoTo Failed
    Select Case CStr(dispatch("action"))
        Case "commit": VTApplyPowerPointNativeEquation sessionId, dispatch
        Case "cancel": VTCancelPowerPointNativeEquation sessionId
        Case Else: Err.Raise 5, "VisualTeX", "Invalid native PowerPoint action."
    End Select
    VTWriteTextAtomic resultPath, "ok"
    Exit Sub
Failed:
    errorText = Err.Description
    errorNumber = Err.Number
    On Error Resume Next
    VTWriteTextAtomic resultPath, "error" & vbLf & CStr(errorNumber) & vbLf & errorText
    On Error GoTo 0
End Sub

Public Sub VTApplyPowerPointNativeEquation(ByVal sessionId As String, ByVal dispatch As Object)
    Dim target As VTPowerPointNativeTarget
    Dim destination As PowerPoint.Shape
    Dim insertion As PowerPoint.TextRange
    Dim pasted As PowerPoint.TextRange
    Dim createdShape As Boolean
    Dim mutationStarted As Boolean
    Dim errorNumber As Long
    Dim errorText As String
    Dim insertedStart As Long
    Dim prefixCount As Long
    Dim suffixCount As Long
    Dim formulaLength As Long
    Dim display As Boolean
    Dim undoCount As Long
    Dim undoLimit As Long

    If NativeApplyBusy Then Err.Raise vbObjectError + 7805, "VisualTeX", "A native PowerPoint insertion is already in progress."
    Set target = VTNativeTarget(sessionId)
    If target.Completed Then Exit Sub
    On Error GoTo Failed
    NativeApplyBusy = True
    VTNativeValidateTarget target
    display = (CStr(dispatch("displayMode")) = "block")
    If display And target.PrefixBreak Then prefixCount = 1
    If display And target.SuffixBreak Then suffixCount = 1
    If CStr(dispatch("leadingParagraph")) <> CStr(prefixCount) Or _
       CStr(dispatch("trailingParagraph")) <> CStr(suffixCount) Then
        Err.Raise vbObjectError + 7804, "VisualTeX", "The native paragraph boundaries do not match the original insertion position."
    End If
    target.OwnerWindow.Activate
    target.OwnerWindow.View.GotoSlide target.TargetSlide.SlideIndex
    If Not target.TableShape Is Nothing Then
        Set destination = target.TargetShape
        mutationStarted = True
        VTWriteCellEquation target, display
        target.Completed = True
        On Error Resume Next
        Set target.TargetShape = Nothing
        Set target.TableShape = Nothing
        Set target.TargetSlide = Nothing
        Set target.OwnerWindow = Nothing
        Set target.Owner = Nothing
        NativeApplyBusy = False
        On Error GoTo 0
        Exit Sub
    End If
    If target.TargetShape Is Nothing Then
        Set destination = target.TargetSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0!, 0!, 240!, 90!)
        createdShape = True
        destination.TextFrame.WordWrap = msoFalse
        destination.TextFrame.AutoSize = ppAutoSizeShapeToFitText
        destination.TextFrame.TextRange.Font.Size = CSng(VTParseInvariantDouble(CStr(dispatch("fontSizePt"))))
        Set insertion = destination.TextFrame.TextRange
        insertedStart = 1
    Else
        Set destination = target.TargetShape
        insertedStart = target.TextStart
        Dim replaceLength As Long
        replaceLength = target.TextLength
        If target.IsEdit Then
            insertedStart = target.ReplaceStart
            replaceLength = target.ReplaceLength
        End If
        Set insertion = destination.TextFrame.TextRange.Characters(insertedStart, replaceLength)
        If insertion.Start <> insertedStart Or insertion.Length <> replaceLength Then
            Err.Raise vbObjectError + 7803, "VisualTeX", "PowerPoint could not restore the exact text insertion range."
        End If
    End If
    insertion.Select
    mutationStarted = True
    Set pasted = insertion.Paste
    If pasted Is Nothing Then Err.Raise vbObjectError + 7806, "VisualTeX", "PowerPoint did not return the inserted native equation."
    formulaLength = pasted.Length - prefixCount - suffixCount
    If formulaLength < 1 Then Err.Raise vbObjectError + 7806, "VisualTeX", "PowerPoint did not insert a mathematical text range."
    If display And Not target.IsEdit Then
        With destination.TextFrame.TextRange.Characters(insertedStart + prefixCount, formulaLength).ParagraphFormat
            .Alignment = ppAlignCenter
            .Bullet.Visible = msoFalse
        End With
    End If
    If createdShape Then
        destination.Left = (target.Owner.PageSetup.SlideWidth - destination.Width) / 2!
        destination.Top = (target.Owner.PageSetup.SlideHeight - destination.Height) / 2!
    End If
    target.Completed = True
    ' A caret/focus failure after a successful paste is not a failed insertion.
    On Error Resume Next
    destination.TextFrame.TextRange.Characters(insertedStart + pasted.Length, 0).Select
    Set target.TargetShape = Nothing
    Set target.TargetSlide = Nothing
    Set target.OwnerWindow = Nothing
    Set target.Owner = Nothing
    NativeApplyBusy = False
    On Error GoTo 0
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If createdShape Then
        destination.Delete
    ElseIf mutationStarted And Not destination Is Nothing Then
        ' A table cell's equation is written in several steps, each its own Undo.
        undoLimit = IIf(target.TableShape Is Nothing, 1, 8)
        Do While StrComp(destination.TextFrame.TextRange.Text, target.OriginalText, vbBinaryCompare) <> 0 And undoCount < undoLimit
            Application.CommandBars.ExecuteMso "Undo"
            undoCount = undoCount + 1
        Loop
        If StrComp(destination.TextFrame.TextRange.Text, target.OriginalText, vbBinaryCompare) <> 0 Then
            target.Completed = True
            errorText = errorText & " The original text could not be restored automatically; inspect the text box before retrying."
        End If
    End If
    NativeApplyBusy = False
    On Error GoTo 0
    Err.Raise errorNumber, "VisualTeX PowerPoint native equation", errorText
End Sub

' PowerPoint turns an equation pasted into a table cell into linear text (its own
' copy of an equation too); only typing one there builds it up. A cell therefore
' gets its equation the way typing does: PowerPoint's linear format of the pasted
' equation, made in a scratch text box off the slide (where pasting keeps math),
' written into the cell and built up by Insert Equation and Professional. Linear
' format keeps no formatting: the app spells letter styles into the characters
' for cells (omml_write prepare_mathml_for_table_cell); colors are not kept.
Private Sub VTWriteCellEquation(ByVal target As VTPowerPointNativeTarget, ByVal display As Boolean)
    Dim cellText As PowerPoint.TextRange
    Dim original As String
    Dim linear As String
    Dim start As Long, length As Long
    Dim leading As Boolean, trailing As Boolean
    Dim at As Long, written As Long
    Dim zone As Object

    linear = VTLinearizePastedEquation(target)
    Set cellText = target.TargetShape.TextFrame.TextRange
    original = cellText.Text
    start = target.TextStart
    length = target.TextLength
    If target.IsEdit Then
        start = target.ReplaceStart
        length = target.ReplaceLength
    ElseIf display Then
        If start > 1 Then leading = (Mid$(original, start - 1, 1) <> vbCr)
        If start + length <= Len(original) Then trailing = (Mid$(original, start + length, 1) <> vbCr)
    End If
    ' Assign, never Delete: Delete also removes a space next to the text.
    cellText.Characters(start, length).Text = IIf(leading, vbCr, "") & linear & IIf(trailing, vbCr, "")
    at = start + IIf(leading, 1, 0)
    ' The linear text as PowerPoint holds it (an assignment can normalize characters).
    written = Len(target.TargetShape.TextFrame.TextRange.Text) - (Len(original) - length) - IIf(leading, 1, 0) - IIf(trailing, 1, 0)
    If written <= 0 Then Err.Raise vbObjectError + 7806, "VisualTeX", "PowerPoint did not write the equation into the table cell."
    Set zone = VTCellZoneAt(target.TargetShape, at, written)
    If zone Is Nothing Then
        target.TargetShape.TextFrame.TextRange.Characters(at, written).Select
        VTExecuteEquationCommand "EquationInsertNew"
        Set zone = VTCellZoneAt(target.TargetShape, at, written)
        If zone Is Nothing Then Err.Raise vbObjectError + 7806, "VisualTeX", "PowerPoint did not make an equation of the table cell's linear text."
    End If
    target.TargetShape.TextFrame.TextRange.Characters(zone.Start, zone.Length).Select
    VTExecuteEquationCommand "EquationProfessional"
    Set zone = VTCellZoneAt(target.TargetShape, at, 1)
    If zone Is Nothing Then Err.Raise vbObjectError + 7806, "VisualTeX", "PowerPoint did not build up the table cell's equation."
    If display And Not target.IsEdit Then
        With target.TargetShape.TextFrame.TextRange.Characters(zone.Start, zone.Length).ParagraphFormat
            .Alignment = ppAlignCenter
            .Bullet.Visible = msoFalse
        End With
    End If
    ' A caret failure after a successful write is not a failed insertion.
    On Error Resume Next
    target.TargetShape.TextFrame.TextRange.Characters(zone.Start + zone.Length, 0).Select
    On Error GoTo 0
End Sub

' The equation on the clipboard, pasted into a scratch text box right of the
' slide, in PowerPoint's own linear format.
Private Function VTLinearizePastedEquation(ByVal target As VTPowerPointNativeTarget) As String
    Dim box As PowerPoint.Shape
    Dim zone As Object
    Dim linear As String
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    Set box = target.TargetSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        target.Owner.PageSetup.SlideWidth + 100!, 0!, 800!, 100!)
    box.TextFrame.TextRange.Paste
    Set zone = VTCellZoneAt(box, 1, Len(box.TextFrame.TextRange.Text))
    If zone Is Nothing Then Err.Raise vbObjectError + 7806, "VisualTeX", "PowerPoint did not create the equation for the table cell."
    box.TextFrame.TextRange.Characters(zone.Start, zone.Length).Select
    VTExecuteEquationCommand "EquationLinearFormat"
    Set zone = VTCellZoneAt(box, 1, Len(box.TextFrame.TextRange.Text))
    If zone Is Nothing Then Err.Raise vbObjectError + 7806, "VisualTeX", "PowerPoint lost the equation while writing its linear format."
    linear = zone.Text
    box.Delete
    Set box = Nothing
    If Len(Trim$(linear)) = 0 Then Err.Raise vbObjectError + 7806, "VisualTeX", "PowerPoint gave the equation no linear format."
    VTLinearizePastedEquation = linear
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If Not box Is Nothing Then box.Delete
    On Error GoTo 0
    Err.Raise errorNumber, "VisualTeX", errorText
End Function

' The math zone of a shape's (or cell's) text overlapping [start, start + length).
Private Function VTCellZoneAt(ByVal shapeItem As PowerPoint.Shape, ByVal start As Long, ByVal length As Long) As Object
    Dim fullText As Object
    Dim zone As Object
    Dim index As Long
    Set fullText = shapeItem.TextFrame2.TextRange
    If length < 1 Then length = 1
    For index = 1 To fullText.MathZones.Count
        Set zone = fullText.MathZones(index, 1)
        If zone.Start < start + length And zone.Start + zone.Length > start Then
            Set VTCellZoneAt = zone
            Exit Function
        End If
    Next index
End Function

' PowerPoint's ribbon commands act on the selection once it has taken it in.
Private Sub VTExecuteEquationCommand(ByVal idMso As String)
    Dim attempt As Long
    Dim enabled As Boolean
    Dim waitUntil As Single
    For attempt = 1 To 20
        DoEvents
        On Error Resume Next
        enabled = Application.CommandBars.GetEnabledMso(idMso)
        If Err.Number <> 0 Then enabled = True
        On Error GoTo 0
        If enabled Then Exit For
        waitUntil = Timer + 0.05
        Do While Timer < waitUntil
            DoEvents
        Loop
    Next attempt
    If Not enabled Then Err.Raise vbObjectError + 7807, "VisualTeX", _
        "PowerPoint's " & idMso & " command is not available for the table cell's equation."
    Application.CommandBars.ExecuteMso idMso
    DoEvents
End Sub

Public Sub VTCancelPowerPointNativeEquation(ByVal sessionId As String)
    On Error Resume Next
    If Not NativeTargets Is Nothing Then NativeTargets.Remove sessionId
    On Error GoTo 0
End Sub
