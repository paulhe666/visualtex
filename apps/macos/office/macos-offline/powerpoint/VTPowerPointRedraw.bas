Attribute VB_Name = "VTPowerPointRedraw"
Option Explicit

' LaTeX redraw and formulas back to LaTeX, as Word's LaTeX redraw group does.
' The text of the text boxes in scope goes to VisualTeX, which finds and
' renders the LaTeX (src-tauri office/presentation_redraw.rs). VisualTeX then
' writes each result through VisualTeX_ApplyPresentationItem, last first, one
' call per formula: an equation in place of its LaTeX, or a picture formula in
' place of a text box that holds only that formula. A restore turns equations
' into $...$ or $$...$$ in their text and picture formulas into text boxes
' holding their LaTeX; VisualTeX reads each equation from PowerPoint's own
' copy of it (VisualTeX_CopyPresentationEquation). Tables, groups and charts
' are left alone.

Private Const PresentationHost As String = "powerpoint"
Private Const RedrawSourceName As String = "/presentation-redraw-source.json"
Private Const RestoreSourceName As String = "/presentation-restore-source.json"
Private Const MaxTextBoxes As Long = 4096
Private Const MaxFormulas As Long = 1000
Private Const MaxSourceCharacters As Long = 2000000
Private Const MaxSizeRuns As Long = 4000
Private Const SourceChunkCharacters As Long = 60000
Private Const DefaultFontSizePt As Double = 18#

Public Sub VisualTeX_RedrawSelectionToOmml()
    VTStartPresentationRedraw "selection", "omml"
End Sub

Public Sub VisualTeX_RedrawSlideToOmml()
    VTStartPresentationRedraw "slide", "omml"
End Sub

Public Sub VisualTeX_RedrawPresentationToOmml()
    VTStartPresentationRedraw "presentation", "omml"
End Sub

Public Sub VisualTeX_RedrawSelectionToImage()
    VTStartPresentationRedraw "selection", "image"
End Sub

Public Sub VisualTeX_RedrawSlideToImage()
    VTStartPresentationRedraw "slide", "image"
End Sub

Public Sub VisualTeX_RedrawPresentationToImage()
    VTStartPresentationRedraw "presentation", "image"
End Sub

Public Sub VisualTeX_RestoreSelectionToLatex()
    VTStartPresentationRestore "selection"
End Sub

Public Sub VisualTeX_RestoreSlideToLatex()
    VTStartPresentationRestore "slide"
End Sub

Public Sub VisualTeX_RestorePresentationToLatex()
    VTStartPresentationRestore "presentation"
End Sub

Private Sub VTStartPresentationRedraw(ByVal scope As String, ByVal outputKind As String)
    Dim targets As Collection
    Dim target As Variant
    Dim shapeItem As PowerPoint.Shape
    Dim fullText As PowerPoint.TextRange
    Dim boxText As String
    Dim textStart As Long
    Dim textLength As Long
    Dim boxCount As Long
    Dim characterCount As Long
    Dim buffer As String
    Dim sessionId As String
    Dim sourcePath As String
    Dim written As Boolean
    Dim stage As String
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo Failed
    stage = "require-presentation"
    VTRequireWritablePowerPointPresentation
    If scope = "presentation" Then
        If MsgBox("VisualTeX will find the LaTeX formulas in the whole presentation and redraw them in place. Continue?", _
           vbOKCancel Or vbQuestion, "VisualTeX") <> vbOK Then Exit Sub
    End If
    stage = "collect-text"
    Set targets = VTPresentationScope(scope)
    sessionId = VTNewUuidV4()
    sourcePath = VTSessionDirectory(sessionId) & RedrawSourceName
    buffer = "{""segments"":["
    For Each target In targets
        Set shapeItem = target(1)
        If VTIsPresentationTextShape(shapeItem) Then
            Set fullText = shapeItem.TextFrame.TextRange
            boxText = fullText.Text
            textStart = 1
            textLength = Len(boxText)
            If target(3) > 0 Then
                textStart = target(2)
                textLength = target(3)
            End If
            If textLength > 0 And textStart >= 1 And textStart + textLength - 1 <= Len(boxText) Then
                boxCount = boxCount + 1
                characterCount = characterCount + textLength
                If boxCount > MaxTextBoxes Or characterCount > MaxSourceCharacters Then
                    Err.Raise vbObjectError + 7950, "VisualTeX", _
                        "There is too much text to redraw at once. Redraw one slide or a selection at a time."
                End If
                If boxCount > 1 Then buffer = buffer & ","
                buffer = buffer & "{""slideId"":" & CStr(target(0).SlideID) & _
                    ",""shapeId"":" & CStr(shapeItem.Id) & _
                    ",""start"":" & CStr(textStart) & _
                    ",""text"":" & VTPresentationJsonText(Mid$(boxText, textStart, textLength)) & _
                    ",""whole"":" & VTJsonBoolean(textStart = 1 And textLength = Len(boxText) And VTIsPlainTextBox(shapeItem)) & _
                    ",""sizes"":" & VTPresentationSizeRuns(fullText, textStart, textLength) & "}"
                If Len(buffer) >= SourceChunkCharacters Then
                    VTWritePresentationSource sourcePath, buffer, written
                    buffer = ""
                End If
            End If
        End If
    Next target
    If boxCount = 0 Then
        If scope = "selection" Then
            VTShowInformation "The selection holds no text of a text box. Select text, text boxes or slides."
        Else
            VTShowInformation "No text box with text was found."
        End If
        Exit Sub
    End If
    VTWritePresentationSource sourcePath, buffer & "]}", written
    stage = "write-request"
    VTWriteRequest sessionId, VTPresentationOperationJson(sessionId, "latexRedraw", scope, outputKind, "")
    stage = "launch"
    VTLaunchSession PresentationHost, sessionId
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Resume ReportFailure
ReportFailure:
    On Error Resume Next
    If written Then VTDeletePresentationSession sessionId
    On Error GoTo 0
    VTShowError "PowerPoint LaTeX redraw (" & stage & ")", errorNumber, errorText
End Sub

Private Sub VTStartPresentationRestore(ByVal scope As String)
    Dim targets As Collection
    Dim target As Variant
    Dim shapeItem As PowerPoint.Shape
    Dim fullText As Object
    Dim zone As Object
    Dim index As Long
    Dim caretZone As Long
    Dim wanted As Boolean
    Dim fontSizePt As Double
    Dim buffer As String
    Dim formulaCount As Long
    Dim equationCount As Long
    Dim sourceKind As String
    Dim sessionId As String
    Dim sourcePath As String
    Dim written As Boolean
    Dim stage As String
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo Failed
    stage = "require-presentation"
    VTRequireWritablePowerPointPresentation
    If scope = "presentation" Then
        If MsgBox("VisualTeX will turn every formula in the presentation into LaTeX code. Continue?", _
           vbOKCancel Or vbQuestion, "VisualTeX") <> vbOK Then Exit Sub
    End If
    stage = "collect-formulas"
    Set targets = VTPresentationScope(scope)
    sessionId = VTNewUuidV4()
    sourcePath = VTSessionDirectory(sessionId) & RestoreSourceName
    buffer = "{""targets"":["
    For Each target In targets
        Set shapeItem = target(1)
        If VTIsVisualTeXPowerPointShape(shapeItem) Then
            formulaCount = formulaCount + 1
            If formulaCount > 1 Then buffer = buffer & ","
            buffer = buffer & "{""slideId"":" & CStr(target(0).SlideID) & _
                ",""shapeId"":" & CStr(shapeItem.Id) & _
                ",""kind"":""image"",""displayMode"":""block""" & _
                ",""fontSizePt"":" & VTJsonNumber(VTPresentationPictureSize(shapeItem)) & _
                ",""metadata"":" & VTPresentationJsonText(VTShapeMetadata(shapeItem)) & "}"
        ElseIf VTIsPresentationTextShape(shapeItem) Then
            Set fullText = shapeItem.TextFrame2.TextRange
            ' The text cursor in an equation: that equation; elsewhere in the
            ' box: all of its equations.
            caretZone = 0
            If target(4) > 0 Then
                For index = 1 To fullText.MathZones.Count
                    Set zone = fullText.MathZones(index, 1)
                    If target(4) >= zone.Start And target(4) < zone.Start + zone.Length Then caretZone = index
                Next index
            End If
            For index = 1 To fullText.MathZones.Count
                Set zone = fullText.MathZones(index, 1)
                If target(3) > 0 Then
                    wanted = (zone.Start < target(2) + target(3) And zone.Start + zone.Length > target(2))
                Else
                    wanted = (caretZone = 0 Or caretZone = index)
                End If
                If wanted And zone.Length > 0 Then
                    formulaCount = formulaCount + 1
                    equationCount = equationCount + 1
                    fontSizePt = 0#
                    On Error Resume Next
                    fontSizePt = zone.Font.Size
                    If fontSizePt < 1# Or fontSizePt > 512# Then fontSizePt = zone.Characters(1, 1).Font.Size
                    On Error GoTo Failed
                    If fontSizePt < 1# Or fontSizePt > 512# Then fontSizePt = DefaultFontSizePt
                    If formulaCount > 1 Then buffer = buffer & ","
                    buffer = buffer & "{""slideId"":" & CStr(target(0).SlideID) & _
                        ",""shapeId"":" & CStr(shapeItem.Id) & _
                        ",""kind"":""omml""" & _
                        ",""start"":" & CStr(zone.Start) & _
                        ",""length"":" & CStr(zone.Length) & _
                        ",""displayMode"":" & IIf(VTPresentationZoneDisplay(fullText, zone), """block""", """inline""") & _
                        ",""fontSizePt"":" & VTJsonNumber(fontSizePt) & "}"
                End If
            Next index
        End If
        If formulaCount > MaxFormulas Then
            Err.Raise vbObjectError + 7950, "VisualTeX", _
                "There are more than 1000 formulas. Turn one slide or a selection at a time into LaTeX."
        End If
        If Len(buffer) >= SourceChunkCharacters Then
            VTWritePresentationSource sourcePath, buffer, written
            buffer = ""
        End If
    Next target
    If formulaCount = 0 Then
        If scope = "selection" Then
            VTShowInformation "No formula was found in the selection."
        Else
            VTShowInformation "No formula was found."
        End If
        Exit Sub
    End If
    VTWritePresentationSource sourcePath, buffer & "]}", written
    If equationCount = formulaCount Then
        sourceKind = "omml"
    ElseIf equationCount = 0 Then
        sourceKind = "image"
    End If
    stage = "write-request"
    VTWriteRequest sessionId, VTPresentationOperationJson(sessionId, "formulaRestore", scope, "latex", sourceKind)
    stage = "launch"
    VTLaunchSession PresentationHost, sessionId
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Resume ReportFailure
ReportFailure:
    On Error Resume Next
    If written Then VTDeletePresentationSession sessionId
    On Error GoTo 0
    VTShowError "PowerPoint formula to LaTeX (" & stage & ")", errorNumber, errorText
End Sub

' VisualTeX reads an equation to restore from PowerPoint's own copy of it
' (presentation_redraw.rs copy_powerpoint_equations); a reason is written
' only when it cannot be copied.
Public Sub VisualTeX_CopyPresentationEquation(ByVal sessionId As String, ByVal indexText As String)
    Dim directory As String
    Dim index As Long
    Dim item As Collection
    Dim targetPresentation As PowerPoint.Presentation
    Dim slideItem As PowerPoint.Slide
    Dim shapeItem As PowerPoint.Shape
    Dim zone As Object
    Dim errorText As String

    On Error GoTo Failed
    directory = VTSessionDirectory(sessionId)
    index = CLng(indexText)
    Set item = VTReadPresentationItem(directory & "/presentation-copy-" & CStr(index) & ".in", sessionId)
    Set shapeItem = VTPresentationItemShape(item, targetPresentation, slideItem)
    Set zone = VTPresentationExactZone(shapeItem, CLng(VTItemValue(item, "start")), CLng(VTItemValue(item, "length")))
    zone.Copy
    Exit Sub

Failed:
    errorText = Err.Description
    Resume ReportFailure
ReportFailure:
    On Error Resume Next
    If Len(directory) > 0 Then
        VTWriteTextAtomic directory & "/presentation-copy-" & CStr(index) & ".txt", "error" & vbLf & errorText
    End If
End Sub

' VisualTeX writes one formula of a redraw or a restore per call, last first:
' the Session is the active one, the formula the one presentation-next.txt
' names. VisualTeX reads the outcome from presentation-item-<index>.txt.
Public Sub VisualTeX_ApplyPresentationItem()
    Dim sessionId As String
    Dim directory As String
    Dim index As Long
    Dim item As Collection
    Dim outcome As String

    On Error GoTo Failed
    sessionId = VTReadActiveSessionId(PresentationHost)
    directory = VTSessionDirectory(sessionId)
    index = CLng(Trim$(VTReadText(directory & "/presentation-next.txt", 32)))
    Set item = VTReadPresentationItem(directory & "/presentation-item-" & CStr(index) & ".in", sessionId)
    VTApplyPresentationItem item
    outcome = "ok"
    GoTo Report

Failed:
    outcome = "error" & vbLf & Err.Description
    Resume Report
Report:
    On Error Resume Next
    If Len(directory) > 0 Then
        VTWriteTextAtomic directory & "/presentation-item-" & CStr(index) & ".txt", outcome
    End If
End Sub

Private Sub VTApplyPresentationItem(ByVal item As Collection)
    Dim targetPresentation As PowerPoint.Presentation
    Dim slideItem As PowerPoint.Slide
    Dim shapeItem As PowerPoint.Shape

    Set shapeItem = VTPresentationItemShape(item, targetPresentation, slideItem)
    Select Case VTItemValue(item, "kind")
        Case "equation"
            VTWriteRedrawEquation targetPresentation, slideItem, shapeItem, item
        Case "picture"
            VTWriteRedrawPicture targetPresentation, slideItem, shapeItem, item
        Case "text"
            If VTItemValue(item, "replace") = "picture" Then
                VTWriteRestoredPicture slideItem, shapeItem, item
            Else
                VTWriteRestoredEquation shapeItem, item
            End If
        Case Else
            Err.Raise vbObjectError + 7951, "VisualTeX", "VisualTeX sent PowerPoint an unknown kind of formula."
    End Select
End Sub

' An equation in place of its LaTeX. A display equation gets a paragraph of
' its own; the text on both sides keeps its paragraph's format.
Private Sub VTWriteRedrawEquation(ByVal targetPresentation As PowerPoint.Presentation, _
    ByVal slideItem As PowerPoint.Slide, ByVal shapeItem As PowerPoint.Shape, ByVal item As Collection)

    Dim fullText As PowerPoint.TextRange
    Dim original As String
    Dim prepared As String
    Dim after As String
    Dim textStart As Long
    Dim textLength As Long
    Dim insertStart As Long
    Dim insertLength As Long
    Dim pastedLength As Long
    Dim display As Boolean
    Dim breakBefore As Boolean
    Dim breakAfter As Boolean
    Dim listItem As Boolean
    Dim changed As Boolean
    Dim insertion As PowerPoint.TextRange
    Dim errorNumber As Long
    Dim errorText As String

    If Not VTIsPresentationTextShape(shapeItem) Then
        Err.Raise vbObjectError + 7952, "VisualTeX", "The text box is no longer a text box."
    End If
    Set fullText = shapeItem.TextFrame.TextRange
    original = fullText.Text
    textStart = CLng(VTItemValue(item, "start"))
    textLength = CLng(VTItemValue(item, "length"))
    VTCheckReadText original, item
    If textStart < 1 Or textLength < 1 Or _
       StrComp(Mid$(original, textStart, textLength), VTItemText(item, "sourceText"), vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 7954, "VisualTeX", _
            "The text box changed after VisualTeX read it; this formula was left as it is."
    End If
    display = (VTItemValue(item, "displayMode") = "block")
    If display Then
        If textStart > 1 Then breakBefore = (Mid$(original, textStart - 1, 1) <> vbCr)
        If textStart + textLength <= Len(original) Then breakAfter = (Mid$(original, textStart + textLength, 1) <> vbCr)
        On Error Resume Next
        listItem = (fullText.Characters(textStart, 1).ParagraphFormat.Bullet.Visible = msoTrue)
        On Error GoTo 0
    End If

    On Error GoTo Failed
    VTShowPresentationSlide targetPresentation, slideItem
    insertStart = textStart
    insertLength = textLength
    If breakBefore Or breakAfter Then
        changed = True
        fullText.Characters(textStart, textLength).Text = IIf(breakBefore, vbCr, "") & IIf(breakAfter, vbCr, "")
        If breakBefore Then insertStart = textStart + 1
        insertLength = 0
    End If
    prepared = shapeItem.TextFrame.TextRange.Text
    Set insertion = shapeItem.TextFrame.TextRange.Characters(insertStart, insertLength)
    If insertion.Start <> insertStart Or insertion.Length <> insertLength Then
        Err.Raise vbObjectError + 7955, "VisualTeX", "PowerPoint could not select the text of the formula."
    End If
    insertion.Select
    changed = True
    insertion.Paste
    after = shapeItem.TextFrame.TextRange.Text
    pastedLength = Len(after) - (Len(prepared) - insertLength)
    If pastedLength <= 0 Then
        Err.Raise vbObjectError + 7956, "VisualTeX", "PowerPoint did not paste the equation."
    End If
    If StrComp(Left$(after, insertStart - 1), Left$(prepared, insertStart - 1), vbBinaryCompare) <> 0 Or _
       StrComp(Mid$(after, insertStart + pastedLength), Mid$(prepared, insertStart + insertLength), vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 7956, "VisualTeX", "PowerPoint did not paste the equation in place of its LaTeX."
    End If
    If Not VTPresentationHasZone(shapeItem, insertStart, pastedLength) Then
        Err.Raise vbObjectError + 7956, "VisualTeX", "PowerPoint pasted the equation as plain text."
    End If
    If display Then
        With shapeItem.TextFrame.TextRange.Characters(insertStart, pastedLength).ParagraphFormat
            .Alignment = ppAlignCenter
            .Bullet.Visible = msoFalse
        End With
        ' In a list item, the text after the equation goes on with the same
        ' item instead of starting a new one.
        If listItem And breakBefore And breakAfter Then
            On Error Resume Next
            shapeItem.TextFrame.TextRange.Characters(insertStart + pastedLength + 1, 1).ParagraphFormat.Bullet.Visible = msoFalse
            On Error GoTo Failed
        End If
    End If
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Resume RestoreText
RestoreText:
    On Error Resume Next
    If changed Then VTUndoUntilText shapeItem, original, 4
    On Error GoTo 0
    Err.Raise errorNumber, "VisualTeX", errorText
End Sub

' A picture formula in place of the text box that held only its LaTeX: at the
' box's centre and z-order, with the identity VTFinalizePowerPointDispatch
' gives a formula.
Private Sub VTWriteRedrawPicture(ByVal targetPresentation As PowerPoint.Presentation, _
    ByVal slideItem As PowerPoint.Slide, ByVal shapeItem As PowerPoint.Shape, ByVal item As Collection)

    Dim formulaId As String
    Dim metadata As String
    Dim formulaReference As String
    Dim imagePath As String
    Dim fallbackPath As String
    Dim widthPt As Double
    Dim heightPt As Double
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim centerX As Double
    Dim centerY As Double
    Dim zOrder As Long
    Dim pictureShape As PowerPoint.Shape
    Dim errorNumber As Long
    Dim errorText As String

    If Not VTIsPresentationTextShape(shapeItem) Or Not VTIsPlainTextBox(shapeItem) Then
        Err.Raise vbObjectError + 7952, "VisualTeX", "Only a text box that holds just the formula becomes a picture formula."
    End If
    If VTItemValue(item, "checkStart") <> "1" Or _
       StrComp(shapeItem.TextFrame.TextRange.Text, VTItemText(item, "checkText"), vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 7954, "VisualTeX", _
            "The text box changed after VisualTeX read it; this formula was left as it is."
    End If
    formulaId = VTItemValue(item, "formulaId")
    metadata = VTItemValue(item, "metadata")
    formulaReference = VTItemValue(item, "formulaReference")
    If Not VTIsCanonicalUuid(formulaId) Or Not VTIsEncodedMetadata(metadata) Then
        Err.Raise vbObjectError + 7957, "VisualTeX", "VisualTeX sent PowerPoint an invalid picture formula."
    End If
    imagePath = VTItemValue(item, "imagePath")
    VTValidateAbsoluteVisualTeXPath imagePath
    If VTCollectionHasKey(item, "fallbackImagePath") Then
        fallbackPath = VTItemValue(item, "fallbackImagePath")
        VTValidateAbsoluteVisualTeXPath fallbackPath
    End If
    widthPt = VTParseInvariantDouble(VTItemValue(item, "widthPt"))
    heightPt = VTParseInvariantDouble(VTItemValue(item, "heightPt"))
    fontSizePt = VTParseInvariantDouble(VTItemValue(item, "fontSizePt"))
    referenceWidthPt = VTParseInvariantDouble(VTItemValue(item, "referenceWidthPt"))
    referenceHeightPt = VTParseInvariantDouble(VTItemValue(item, "referenceHeightPt"))
    If widthPt <= 0# Or heightPt <= 0# Then
        Err.Raise vbObjectError + 7957, "VisualTeX", "VisualTeX sent PowerPoint a picture formula without a size."
    End If
    centerX = shapeItem.Left + shapeItem.Width / 2#
    centerY = shapeItem.Top + shapeItem.Height / 2#
    zOrder = shapeItem.ZOrderPosition

    On Error Resume Next
    Set pictureShape = slideItem.Shapes.AddPicture( _
        FileName:=imagePath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=CSng(centerX - widthPt / 2#), Top:=CSng(centerY - heightPt / 2#), _
        Width:=CSng(widthPt), Height:=CSng(heightPt))
    Err.Clear
    On Error GoTo Failed
    If pictureShape Is Nothing And Len(fallbackPath) > 0 Then
        Set pictureShape = slideItem.Shapes.AddPicture( _
            FileName:=fallbackPath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
            Left:=CSng(centerX - widthPt / 2#), Top:=CSng(centerY - heightPt / 2#), _
            Width:=CSng(widthPt), Height:=CSng(heightPt))
    End If
    If pictureShape Is Nothing Then
        Err.Raise vbObjectError + 7957, "VisualTeX", "PowerPoint could not insert the formula picture."
    End If
    pictureShape.LockAspectRatio = msoTrue
    If Abs(shapeItem.Rotation) > 0.01 Then pictureShape.Rotation = shapeItem.Rotation
    pictureShape.Name = "VisualTeX_" & formulaId
    pictureShape.AlternativeText = metadata
    pictureShape.Title = formulaReference
    VTSetShapeTag pictureShape, "VisualTeXFormulaId", formulaId
    VTSetShapeTag pictureShape, "VisualTeXDocumentId", VTPresentationIdentityFor(targetPresentation)
    VTSetShapeTag pictureShape, "VisualTeXObjectId", VTPowerPointShapeObjectIdentity(slideItem, pictureShape)
    VTSetShapeTag pictureShape, "VisualTeXSessionId", VTItemValue(item, "sessionId")
    VTSetShapeTag pictureShape, "VisualTeXPending", "0"
    VTSetShapeTag pictureShape, "VisualTeXFormulaReference", formulaReference
    VTSetPowerPointFormulaScaleState pictureShape, fontSizePt, referenceWidthPt, referenceHeightPt
    On Error Resume Next
    VTSetShapeTag pictureShape, "VisualTeXMetadata", metadata
    Err.Clear
    On Error GoTo Failed
    VTRestoreZOrder pictureShape, zOrder + 1
    shapeItem.Delete
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Resume RemovePicture
RemovePicture:
    On Error Resume Next
    If Not pictureShape Is Nothing Then pictureShape.Delete
    On Error GoTo 0
    Err.Raise errorNumber, "VisualTeX", errorText
End Sub

' The LaTeX in place of an equation, as ordinary text.
Private Sub VTWriteRestoredEquation(ByVal shapeItem As PowerPoint.Shape, ByVal item As Collection)
    Dim zoneStart As Long
    Dim zoneLength As Long
    Dim latex As String
    Dim original As String
    Dim fontName As String
    Dim writtenText As PowerPoint.TextRange
    Dim changed As Boolean
    Dim errorNumber As Long
    Dim errorText As String

    zoneStart = CLng(VTItemValue(item, "start"))
    zoneLength = CLng(VTItemValue(item, "length"))
    Call VTPresentationExactZone(shapeItem, zoneStart, zoneLength)
    latex = VTItemText(item, "text")
    original = shapeItem.TextFrame.TextRange.Text
    ' The font of the text next to the equation (an equation's own font is
    ' Cambria Math).
    On Error Resume Next
    If zoneStart > 1 Then
        fontName = shapeItem.TextFrame.TextRange.Characters(zoneStart - 1, 1).Font.Name
    ElseIf zoneStart + zoneLength <= Len(original) Then
        fontName = shapeItem.TextFrame.TextRange.Characters(zoneStart + zoneLength, 1).Font.Name
    End If
    Err.Clear

    On Error GoTo Failed
    changed = True
    ' Assign, never Delete: Delete also removes a space next to the text.
    shapeItem.TextFrame.TextRange.Characters(zoneStart, zoneLength).Text = latex
    Set writtenText = shapeItem.TextFrame.TextRange.Characters(zoneStart, Len(latex))
    If VTPresentationHasZone(shapeItem, zoneStart, Len(latex)) Then
        Err.Raise vbObjectError + 7958, "VisualTeX", "PowerPoint kept the LaTeX inside an equation."
    End If
    If Len(fontName) > 0 And StrComp(fontName, "Cambria Math", vbTextCompare) <> 0 Then
        On Error Resume Next
        If StrComp(writtenText.Font.Name, fontName, vbBinaryCompare) <> 0 Then writtenText.Font.Name = fontName
        Err.Clear
    End If
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Resume RestoreText
RestoreText:
    On Error Resume Next
    If changed Then VTUndoUntilText shapeItem, original, 2
    On Error GoTo 0
    Err.Raise errorNumber, "VisualTeX", errorText
End Sub

' A text box holding the LaTeX in place of a picture formula, at its centre
' and z-order.
Private Sub VTWriteRestoredPicture(ByVal slideItem As PowerPoint.Slide, ByVal shapeItem As PowerPoint.Shape, _
    ByVal item As Collection)

    Dim latex As String
    Dim fontSizePt As Double
    Dim centerX As Double
    Dim centerY As Double
    Dim zOrder As Long
    Dim box As PowerPoint.Shape
    Dim errorNumber As Long
    Dim errorText As String

    If Not VTIsVisualTeXPowerPointShape(shapeItem) Then
        Err.Raise vbObjectError + 7954, "VisualTeX", "The formula is no longer a VisualTeX formula; it was left as it is."
    End If
    If StrComp(VTShapeMetadata(shapeItem), VTItemValue(item, "metadata"), vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 7954, "VisualTeX", "The formula changed after VisualTeX read it; it was left as it is."
    End If
    latex = VTItemText(item, "text")
    fontSizePt = VTParseInvariantDouble(VTItemValue(item, "fontSizePt"))
    centerX = shapeItem.Left + shapeItem.Width / 2#
    centerY = shapeItem.Top + shapeItem.Height / 2#
    zOrder = shapeItem.ZOrderPosition

    On Error GoTo Failed
    Set box = slideItem.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        CSng(centerX - 120#), CSng(centerY - 20#), 240!, 40!)
    box.TextFrame.WordWrap = msoFalse
    box.TextFrame.AutoSize = ppAutoSizeShapeToFitText
    box.TextFrame.TextRange.Text = latex
    If fontSizePt >= 1# And fontSizePt <= 512# Then box.TextFrame.TextRange.Font.Size = CSng(fontSizePt)
    box.Left = CSng(centerX - box.Width / 2#)
    box.Top = CSng(centerY - box.Height / 2#)
    VTRestoreZOrder box, zOrder + 1
    shapeItem.Delete
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Resume RemoveBox
RemoveBox:
    On Error Resume Next
    If Not box Is Nothing Then box.Delete
    On Error GoTo 0
    Err.Raise errorNumber, "VisualTeX", errorText
End Sub

' What a command works on: Array(slide, shape, start, length, caret) for each
' shape. start and length name selected text (length 0: the whole text);
' caret is the text cursor's position (0: none). The selected text, shapes or
' slide thumbnails ("selection"), the slide in view ("slide") or every slide.
Private Function VTPresentationScope(ByVal scope As String) As Collection
    Dim result As New Collection
    Dim slideList As New Collection
    Dim selected As PowerPoint.Selection
    Dim slideItem As PowerPoint.Slide
    Dim shapeItem As PowerPoint.Shape
    Dim index As Long

    If scope = "presentation" Then
        For index = 1 To ActivePresentation.Slides.Count
            slideList.Add ActivePresentation.Slides(index)
        Next index
    ElseIf scope = "selection" Then
        Set selected = ActiveWindow.Selection
        If selected.Type = ppSelectionText Then
            Set shapeItem = selected.ShapeRange(1)
            If selected.TextRange.Length > 0 Then
                result.Add Array(shapeItem.Parent, shapeItem, selected.TextRange.Start, selected.TextRange.Length, 0)
            Else
                result.Add Array(shapeItem.Parent, shapeItem, 0, 0, selected.TextRange.Start)
            End If
            Set VTPresentationScope = result
            Exit Function
        ElseIf selected.Type = ppSelectionShapes Then
            For index = 1 To selected.ShapeRange.Count
                Set shapeItem = selected.ShapeRange(index)
                result.Add Array(shapeItem.Parent, shapeItem, 0, 0, 0)
            Next index
            Set VTPresentationScope = result
            Exit Function
        ElseIf selected.Type = ppSelectionSlides Then
            For index = 1 To selected.SlideRange.Count
                slideList.Add selected.SlideRange(index)
            Next index
        End If
    End If
    If slideList.Count = 0 Then slideList.Add ActiveWindow.View.Slide
    For index = 1 To slideList.Count
        Set slideItem = slideList(index)
        For Each shapeItem In slideItem.Shapes
            result.Add Array(slideItem, shapeItem, 0, 0, 0)
        Next shapeItem
    Next index
    Set VTPresentationScope = result
End Function

' A text box, a placeholder, or a shape drawn as nothing but its text: only
' these become a picture formula (a filled or outlined shape keeps its LaTeX).
Private Function VTIsPlainTextBox(ByVal shapeItem As PowerPoint.Shape) As Boolean
    On Error GoTo NotPlain
    Select Case shapeItem.Type
        Case msoTextBox, msoPlaceholder
            VTIsPlainTextBox = True
        Case msoAutoShape
            VTIsPlainTextBox = (shapeItem.Fill.Visible <> msoTrue And shapeItem.Line.Visible <> msoTrue)
    End Select
    Exit Function
NotPlain:
    VTIsPlainTextBox = False
End Function

' A shape whose own text VisualTeX reads and writes: not a table, group,
' chart or VisualTeX picture formula.
Private Function VTIsPresentationTextShape(ByVal shapeItem As PowerPoint.Shape) As Boolean
    On Error GoTo NotText
    If shapeItem.Type = msoGroup Then Exit Function
    If shapeItem.HasTable = msoTrue Then Exit Function
    If shapeItem.HasTextFrame <> msoTrue Then Exit Function
    If VTIsVisualTeXPowerPointShape(shapeItem) Then Exit Function
    VTIsPresentationTextShape = True
    Exit Function
NotText:
    VTIsPresentationTextShape = False
End Function

' JSON text for a text box's text; Replace$ keeps a long text fast.
Private Function VTPresentationJsonText(ByVal value As String) As String
    Dim code As Long

    value = Replace$(value, "\", "\\")
    value = Replace$(value, """", "\""")
    value = Replace$(value, vbCr, "\r")
    value = Replace$(value, vbLf, "\n")
    value = Replace$(value, vbTab, "\t")
    For code = 0 To 31
        If code <> 9 And code <> 10 And code <> 13 Then
            If InStr(1, value, ChrW(code), vbBinaryCompare) > 0 Then
                value = Replace$(value, ChrW(code), "\u" & Right$("000" & Hex$(code), 4))
            End If
        End If
    Next code
    VTPresentationJsonText = """" & value & """"
End Function

' The font sizes over a text as [start, length, size] runs, 0-based in it.
Private Function VTPresentationSizeRuns(ByVal fullText As PowerPoint.TextRange, _
    ByVal textStart As Long, ByVal textLength As Long) As String

    Dim part As PowerPoint.TextRange
    Dim runRange As PowerPoint.TextRange
    Dim index As Long
    Dim runStart As Long
    Dim runLength As Long
    Dim size As Double
    Dim result As String

    On Error GoTo RunsFailed
    Set part = fullText.Characters(textStart, textLength)
    For index = 1 To part.Runs.Count
        If index > MaxSizeRuns Then Exit For
        Set runRange = part.Runs(index, 1)
        runStart = runRange.Start - textStart
        runLength = runRange.Length
        If runStart < 0 Then
            runLength = runLength + runStart
            runStart = 0
        End If
        If runStart + runLength > textLength Then runLength = textLength - runStart
        size = runRange.Font.Size
        If runLength > 0 And size >= 1# And size <= 4000# Then
            If Len(result) > 0 Then result = result & ","
            result = result & "[" & CStr(runStart) & "," & CStr(runLength) & "," & VTJsonNumber(size) & "]"
        End If
    Next index
    If Len(result) > 0 Then
        VTPresentationSizeRuns = "[" & result & "]"
        Exit Function
    End If
    GoTo FirstSize

RunsFailed:
    Resume FirstSize
FirstSize:
    On Error Resume Next
    size = 0#
    size = fullText.Characters(textStart, 1).Font.Size
    On Error GoTo 0
    If size < 1# Or size > 4000# Then size = DefaultFontSizePt
    VTPresentationSizeRuns = "[[0," & CStr(textLength) & "," & VTJsonNumber(size) & "]]"
End Function

' A display equation is alone in its paragraph (as VTNativeEditRequest decides).
Private Function VTPresentationZoneDisplay(ByVal fullText As Object, ByVal zone As Object) As Boolean
    Dim paragraph As Object
    Dim index As Long
    Dim beforeLength As Long
    Dim afterLength As Long
    Dim outside As String

    For index = 1 To fullText.Paragraphs.Count
        Set paragraph = fullText.Paragraphs(index, 1)
        If paragraph.Start <= zone.Start And paragraph.Start + paragraph.Length >= zone.Start + zone.Length Then
            beforeLength = zone.Start - paragraph.Start
            afterLength = paragraph.Start + paragraph.Length - zone.Start - zone.Length
            If beforeLength > 0 Then outside = fullText.Characters(paragraph.Start, beforeLength).Text
            If afterLength > 0 Then outside = outside & fullText.Characters(zone.Start + zone.Length, afterLength).Text
            outside = Replace$(Replace$(outside, vbCr, ""), Chr$(11), "")
            VTPresentationZoneDisplay = (Len(Trim$(outside)) = 0)
            Exit Function
        End If
    Next index
End Function

' A picture formula's size, from its tag (read only).
Private Function VTPresentationPictureSize(ByVal shapeItem As PowerPoint.Shape) As Double
    Dim value As Double

    On Error Resume Next
    value = VTParseInvariantDouble(shapeItem.Tags("VisualTeXFontSizePt"))
    On Error GoTo 0
    If value < 1# Or value > 512# Then value = DefaultFontSizePt
    VTPresentationPictureSize = value
End Function

' The request as Word's document operations write it, for PowerPoint.
Private Function VTPresentationOperationJson(ByVal sessionId As String, ByVal operationName As String, _
    ByVal scope As String, ByVal outputKind As String, ByVal sourceKind As String) As String

    Dim options As String

    options = """bookmarkName"":" & VTJsonString("VT_D_" & Replace$(Left$(sessionId, 24), "-", "")) & _
        ",""defaultFontSizePt"":" & VTJsonNumber(DefaultFontSizePt) & _
        ",""redrawScope"":" & VTJsonString(IIf(scope = "presentation", "document", "selection")) & _
        ",""outputKind"":" & VTJsonString(outputKind)
    If Len(sourceKind) > 0 Then options = options & ",""sourceKind"":" & VTJsonString(sourceKind)
    VTPresentationOperationJson = "{" & _
        """protocolVersion"":" & CStr(VT_PROTOCOL_VERSION) & "," & _
        """sessionId"":" & VTJsonString(sessionId) & "," & _
        """host"":" & VTJsonString(PresentationHost) & ",""mode"":""create""," & _
        """operation"":" & VTJsonString(operationName) & "," & _
        """formulaId"":null,""displayMode"":""inline"",""numbered"":false,""nativeEquation"":false," & _
        """sourceDocumentId"":" & VTJsonString(VTPresentationIdentityFor(ActivePresentation)) & "," & _
        """sourceObjectId"":null,""encodedMetadata"":null,""pendingMarker"":null," & _
        """fontSizePt"":null,""referenceWidthPt"":null,""referenceHeightPt"":null,""powerPoint"":null," & _
        """documentImport"":{" & options & "}}"
End Function

' The source goes to VisualTeX in parts the file bridge takes.
Private Sub VTWritePresentationSource(ByVal sourcePath As String, ByVal contents As String, ByRef written As Boolean)
    Dim offset As Long
    Dim part As String
    Dim lastCode As Long

    offset = 1
    Do While offset <= Len(contents)
        part = Mid$(contents, offset, SourceChunkCharacters)
        If Len(part) > 1 And offset + Len(part) <= Len(contents) Then
            ' Never split a surrogate pair between two parts.
            lastCode = AscW(Right$(part, 1))
            If lastCode >= -10240 And lastCode <= -9217 Then part = Left$(part, Len(part) - 1)
        End If
        If written Then
            VTAppendText sourcePath, part
        Else
            VTWriteTextAtomic sourcePath, part
            written = True
        End If
        offset = offset + Len(part)
    Loop
End Sub

Private Sub VTDeletePresentationSession(ByVal sessionId As String)
    On Error Resume Next
    Call VTFileBridgeCall("DeleteVisualTeXFile", "OfficeSessions/" & sessionId & RedrawSourceName)
    Call VTFileBridgeCall("DeleteVisualTeXFile", "OfficeSessions/" & sessionId & RestoreSourceName)
    VTDeleteSessionFiles sessionId
    On Error GoTo 0
End Sub

' A key=value item VisualTeX wrote for this Session (presentation_redraw.rs).
Private Function VTReadPresentationItem(ByVal itemPath As String, ByVal sessionId As String) As Collection
    Dim result As New Collection
    Dim rows() As String
    Dim index As Long
    Dim separator As Long
    Dim key As String

    rows = Split(Replace$(VTReadText(itemPath, 4194304), vbCrLf, vbLf), vbLf)
    For index = LBound(rows) To UBound(rows)
        If Len(rows(index)) > 0 Then
            separator = InStr(1, rows(index), "=", vbBinaryCompare)
            If separator <= 1 Then
                Err.Raise vbObjectError + 7951, "VisualTeX", "VisualTeX sent PowerPoint an invalid formula item."
            End If
            key = Left$(rows(index), separator - 1)
            If VTCollectionHasKey(result, key) Then
                Err.Raise vbObjectError + 7951, "VisualTeX", "VisualTeX sent PowerPoint a formula item with " & key & " twice."
            End If
            result.Add Mid$(rows(index), separator + 1), key
        End If
    Next index
    If VTItemValue(result, "protocolVersion") <> CStr(VT_PROTOCOL_VERSION) Or _
       VTItemValue(result, "sessionId") <> sessionId Then
        Err.Raise vbObjectError + 7951, "VisualTeX", "The formula item belongs to another VisualTeX operation."
    End If
    Set VTReadPresentationItem = result
End Function

Private Function VTItemValue(ByVal item As Collection, ByVal key As String) As String
    If Not VTCollectionHasKey(item, key) Then
        Err.Raise vbObjectError + 7951, "VisualTeX", "VisualTeX sent PowerPoint a formula item without " & key & "."
    End If
    VTItemValue = CStr(item(key))
End Function

Private Function VTItemText(ByVal item As Collection, ByVal key As String) As String
    VTItemText = VTBase64UrlDecodeUtf8(VTItemValue(item, key))
End Function

' The shape an item names, in the presentation it came from.
Private Function VTPresentationItemShape(ByVal item As Collection, _
    ByRef targetPresentation As PowerPoint.Presentation, ByRef slideItem As PowerPoint.Slide) As PowerPoint.Shape

    Dim slideId As Long
    Dim shapeId As Long

    slideId = CLng(VTItemValue(item, "slideId"))
    shapeId = CLng(VTItemValue(item, "shapeId"))
    Set targetPresentation = VTFindPowerPointPresentation(VTItemValue(item, "presentationIdentity"))
    VTRequireWritablePowerPointPresentationObject targetPresentation
    Set slideItem = VTPresentationSlideById(targetPresentation, slideId)
    Set VTPresentationItemShape = VTFindPowerPointShapeById(slideItem, shapeId)
    If VTPresentationItemShape Is Nothing Then
        Err.Raise vbObjectError + 7952, "VisualTeX", "The text box or formula is no longer on its slide."
    End If
End Function

Private Function VTPresentationSlideById(ByVal targetPresentation As PowerPoint.Presentation, _
    ByVal slideId As Long) As PowerPoint.Slide

    On Error Resume Next
    Set VTPresentationSlideById = targetPresentation.Slides.FindBySlideID(slideId)
    On Error GoTo 0
    If VTPresentationSlideById Is Nothing Then
        Err.Raise vbObjectError + 7952, "VisualTeX", "The slide is no longer in the presentation."
    End If
End Function

' The equation exactly at [start, start + length) of a shape's text.
Private Function VTPresentationExactZone(ByVal shapeItem As PowerPoint.Shape, _
    ByVal zoneStart As Long, ByVal zoneLength As Long) As Object

    Dim fullText As Object
    Dim zone As Object
    Dim index As Long

    If VTIsPresentationTextShape(shapeItem) Then
        Set fullText = shapeItem.TextFrame2.TextRange
        For index = 1 To fullText.MathZones.Count
            Set zone = fullText.MathZones(index, 1)
            If zone.Start = zoneStart And zone.Length = zoneLength Then
                Set VTPresentationExactZone = zone
                Exit Function
            End If
        Next index
    End If
    Err.Raise vbObjectError + 7953, "VisualTeX", _
        "The equation changed or moved after VisualTeX read it; it was left as it is."
End Function

' Whether an equation overlaps [start, start + length) of a shape's text.
Private Function VTPresentationHasZone(ByVal shapeItem As PowerPoint.Shape, _
    ByVal textStart As Long, ByVal textLength As Long) As Boolean

    Dim fullText As Object
    Dim zone As Object
    Dim index As Long

    Set fullText = shapeItem.TextFrame2.TextRange
    For index = 1 To fullText.MathZones.Count
        Set zone = fullText.MathZones(index, 1)
        If zone.Length > 0 And zone.Start < textStart + textLength And zone.Start + zone.Length > textStart Then
            VTPresentationHasZone = True
            Exit Function
        End If
    Next index
End Function

' The box's text must still be what VisualTeX read, up to the formula's end
' (the formulas after it are already written).
Private Sub VTCheckReadText(ByVal boxText As String, ByVal item As Collection)
    Dim checkStart As Long
    Dim expected As String

    checkStart = CLng(VTItemValue(item, "checkStart"))
    expected = VTItemText(item, "checkText")
    If checkStart < 1 Or StrComp(Mid$(boxText, checkStart, Len(expected)), expected, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 7954, "VisualTeX", _
            "The text box changed after VisualTeX read it; this formula was left as it is."
    End If
End Sub

' Text can be selected on the slide in view only.
Private Sub VTShowPresentationSlide(ByVal targetPresentation As PowerPoint.Presentation, _
    ByVal slideItem As PowerPoint.Slide)

    Dim windowItem As PowerPoint.DocumentWindow

    Set windowItem = targetPresentation.Windows(1)
    windowItem.Activate
    If windowItem.View.Slide.SlideID <> slideItem.SlideID Then windowItem.View.GotoSlide slideItem.SlideIndex
End Sub

' Undo until the text box holds its original text again, at most `limit` steps.
Private Sub VTUndoUntilText(ByVal shapeItem As PowerPoint.Shape, ByVal original As String, ByVal limit As Long)
    Dim steps As Long

    Do While StrComp(shapeItem.TextFrame.TextRange.Text, original, vbBinaryCompare) <> 0 And steps < limit
        Application.CommandBars.ExecuteMso "Undo"
        DoEvents
        steps = steps + 1
    Loop
End Sub
