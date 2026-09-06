Attribute VB_Name = "VTPowerPointAdapter"
Option Explicit

Private Const VT_POWERPOINT_HOST As String = "powerpoint"
Private Const VT_POWERPOINT_STATUS_FILE As String = "/OfficePluginStatus/powerpoint.json"
Private Const VT_POWERPOINT_SOURCE_REVISION As String = _
    "powerpoint-office-performance-20260801-r4"
Private Const VT_SHAPE_PREFIX As String = "VisualTeX_"
Private Const VT_DEFAULT_PLACEHOLDER_WIDTH As Single = 180!
Private Const VT_DEFAULT_PLACEHOLDER_HEIGHT As Single = 42!
Private Const VT_POWERPOINT_REFERENCE_FONT_SIZE_PT As Double = 14#
Private Const VT_POWERPOINT_DEFAULT_FONT_SIZE_PT As Double = 18#
Private Const VT_POWERPOINT_MIN_FONT_SIZE_PT As Double = 1#
Private Const VT_POWERPOINT_MAX_FONT_SIZE_PT As Double = 512#
Private Const VT_POWERPOINT_FONT_CONTROL_ID As String = _
    "VisualTeX.Mac.PowerPoint.FormulaFontSize"
Private Const VT_POWERPOINT_FONT_TAG As String = "VisualTeXFontSizePt"
Private Const VT_POWERPOINT_DOCUMENT_TAG As String = "VisualTeXDocumentId"
Private Const VT_POWERPOINT_OBJECT_TAG As String = "VisualTeXObjectId"
Private Const VT_POWERPOINT_REFERENCE_WIDTH_TAG As String = _
    "VisualTeXReferenceWidthPt"
Private Const VT_POWERPOINT_REFERENCE_HEIGHT_TAG As String = _
    "VisualTeXReferenceHeightPt"
Private VT_POWERPOINT_EVENT_SINK As VTPowerPointEvents
Private VT_POWERPOINT_RIBBON As IRibbonUI
Private VT_POWERPOINT_SELECTION_SYNC_ACTIVE As Boolean
Private VT_MOVED_COPY_TEST_SOURCE_ID As Long
Private VT_MOVED_COPY_TEST_COPIED_ID As Long
Private VT_MOVED_COPY_TEST_SOURCE_LEFT As Double
Private VT_MOVED_COPY_TEST_SOURCE_TOP As Double
Private VT_MOVED_COPY_TEST_COPIED_LEFT As Double
Private VT_MOVED_COPY_TEST_COPIED_TOP As Double
Private VT_MOVED_COPY_TEST_SOURCE_FORMULA_ID As String

Public Sub Auto_Open()
    On Error Resume Next
    VTInitializePowerPointEvents
    VTPrewarmApplication VT_POWERPOINT_HOST
    VTWritePowerPointHealth
    On Error GoTo 0
End Sub

Public Sub VTInitializePowerPointEvents()
    Set VT_POWERPOINT_EVENT_SINK = New VTPowerPointEvents
    Set VT_POWERPOINT_EVENT_SINK.App = PowerPoint.Application
End Sub

Public Sub VTPowerPointRibbonOnLoad(ByVal ribbon As IRibbonUI)
    Set VT_POWERPOINT_RIBBON = ribbon
    VTInitializePowerPointEvents
    VTInvalidatePowerPointFormulaFontSizeControl
    VTWritePowerPointHealth
End Sub

Public Sub VisualTeX_NewFormula()
    On Error GoTo Failed

    Dim sessionId As String
    Dim formulaId As String
    Dim pendingMarker As String
    Dim currentSlide As Slide
    Dim placeholder As Shape
    Dim slideWidth As Single
    Dim slideHeight As Single
    Dim requestJson As String
    Dim powerPointJson As String
    Dim requestedFontSizePt As Double
    Dim failureStage As String

    failureStage = "validate presentation"
    VTRequireWritablePowerPointPresentation
    Set currentSlide = ActiveWindow.View.Slide

    failureStage = "resolve formula font size"
    requestedFontSizePt = VTPreferredPowerPointFormulaFontSize()

    failureStage = "create identifiers"
    sessionId = VTNewUuidV4()
    formulaId = VTNewUuidV4()
    pendingMarker = VTPendingMarker(sessionId, formulaId)
    slideWidth = ActivePresentation.PageSetup.SlideWidth
    slideHeight = ActivePresentation.PageSetup.SlideHeight

    failureStage = "create placeholder shape"
    Set placeholder = currentSlide.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        (slideWidth - VT_DEFAULT_PLACEHOLDER_WIDTH) / 2!, _
        (slideHeight - VT_DEFAULT_PLACEHOLDER_HEIGHT) / 2!, _
        VT_DEFAULT_PLACEHOLDER_WIDTH, _
        VT_DEFAULT_PLACEHOLDER_HEIGHT)

    failureStage = "format placeholder shape"
    placeholder.Name = VT_SHAPE_PREFIX & formulaId
    placeholder.Fill.Visible = msoFalse
    placeholder.Line.Visible = msoTrue
    placeholder.Line.Weight = 1!
    placeholder.Line.ForeColor.RGB = RGB(128, 128, 128)
    placeholder.TextFrame.TextRange.Text = "VisualTeX"
    placeholder.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter

    failureStage = "attach placeholder metadata"
    VTSetShapeTag placeholder, "VisualTeXFormulaId", formulaId
    VTSetShapeTag placeholder, VT_POWERPOINT_DOCUMENT_TAG, VTPresentationIdentity()
    VTSetShapeTag placeholder, VT_POWERPOINT_OBJECT_TAG, _
        VTPowerPointShapeObjectIdentity(currentSlide, placeholder)
    VTSetShapeTag placeholder, "VisualTeXSessionId", sessionId
    VTSetShapeTag placeholder, "VisualTeXPending", "1"
    VTSetShapeTag placeholder, VT_POWERPOINT_FONT_TAG, _
        VTJsonNumber(requestedFontSizePt)
    placeholder.AlternativeText = pendingMarker

    failureStage = "build request"
    powerPointJson = VTPowerPointGeometryJson( _
        currentSlide, placeholder, requestedFontSizePt)
    requestJson = VTRequestJson( _
        sessionId, _
        VT_POWERPOINT_HOST, _
        "create", _
        formulaId, _
        "block", _
        False, _
        VTPresentationIdentity(), _
        placeholder.Name, _
        "", _
        pendingMarker, _
        powerPointJson)

    failureStage = "write request and open VisualTeX editor"
    Call VTWriteAndLaunchSession( _
        VT_POWERPOINT_HOST, sessionId, requestJson)
    Exit Sub

Failed:
    Dim errorNumber As Long
    Dim errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    If Len(failureStage) > 0 Then
        errorDescription = "Stage: " & failureStage & ". " & errorDescription
    End If
    On Error Resume Next
    If Not placeholder Is Nothing Then placeholder.Delete
    If Len(sessionId) > 0 Then VTDeleteSessionFiles sessionId
    On Error GoTo 0
    VTShowError "PowerPoint formula creation", errorNumber, errorDescription
End Sub

Public Sub VisualTeX_EditSelected()
    On Error GoTo Failed

    VTRequireWritablePowerPointPresentation
    VTPowerPointEditShape VTSelectedSingleShape()
    Exit Sub

Failed:
    VTShowError "PowerPoint edit", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_DoubleClickEditSelected()
    ' Invoked by the native macOS double-click monitor. Let VBA errors return
    ' to the monitor instead of displaying a modal message for ordinary shapes.
    VTRequireWritablePowerPointPresentation
    VTPowerPointEditShape VTSelectedSingleShape()
End Sub

Public Sub VisualTeX_EditShape(ByVal selectedShape As Shape)
    On Error GoTo Failed
    VTRequireWritablePowerPointPresentation
    VTPowerPointEditShape selectedShape
    Exit Sub
Failed:
    VTShowError "PowerPoint edit", Err.Number, Err.Description
End Sub

Public Function VTIsVisualTeXPowerPointShape(ByVal selectedShape As Shape) As Boolean
    Dim formulaId As String
    Dim encodedMetadata As String
    Dim parsedFormulaId As String
    Dim displayMode As String
    Dim numbered As Boolean

    On Error GoTo InvalidShape
    If selectedShape Is Nothing Then Exit Function
    formulaId = VTShapeFormulaId(selectedShape)
    encodedMetadata = VTShapeMetadata(selectedShape)
    If Not VTIsCanonicalUuid(formulaId) Or Not VTIsEncodedMetadata(encodedMetadata) Then Exit Function
    If Not VTTryParseFormulaReference(selectedShape.Title, parsedFormulaId, displayMode, numbered) Then Exit Function
    If parsedFormulaId <> formulaId Then Exit Function
    VTIsVisualTeXPowerPointShape = True
    Exit Function
InvalidShape:
    VTIsVisualTeXPowerPointShape = False
End Function

Private Function VTPowerPointShapeObjectIdentity( _
    ByVal currentSlide As Slide, _
    ByVal target As Shape) As String

    If currentSlide Is Nothing Or target Is Nothing Then Exit Function
    VTPowerPointShapeObjectIdentity = _
        CStr(currentSlide.SlideID) & ":" & CStr(target.Id)
End Function

Private Function VTPowerPointShapeCollectionIndex( _
    ByVal currentSlide As Slide, _
    ByVal target As Shape) As Long

    Dim shapeIndex As Long

    If currentSlide Is Nothing Or target Is Nothing Then Exit Function
    For shapeIndex = 1 To currentSlide.Shapes.Count
        If currentSlide.Shapes(shapeIndex).Id = target.Id Then
            VTPowerPointShapeCollectionIndex = shapeIndex
            Exit Function
        End If
    Next shapeIndex
End Function

Private Function VTFindPowerPointShapeById( _
    ByVal currentSlide As Slide, _
    ByVal shapeId As Long) As Shape

    Dim candidate As Shape

    If currentSlide Is Nothing Or shapeId <= 0 Then Exit Function
    For Each candidate In currentSlide.Shapes
        If candidate.Id = shapeId Then
            Set VTFindPowerPointShapeById = candidate
            Exit Function
        End If
    Next candidate
End Function

Private Function VTShouldForkPowerPointFormulaCopy( _
    ByVal currentSlide As Slide, _
    ByVal selectedShape As Shape, _
    ByVal formulaId As String) As Boolean

    Dim slideObject As Slide
    Dim candidate As Shape
    Dim canonicalSlide As Slide
    Dim canonicalShape As Shape
    Dim matchCount As Long
    Dim ownerDocumentId As String
    Dim ownerObjectId As String
    Dim currentDocumentId As String
    Dim currentObjectId As String

    If currentSlide Is Nothing Or selectedShape Is Nothing Or _
       Not VTIsCanonicalUuid(formulaId) Then Exit Function
    currentDocumentId = VTPresentationIdentity()
    currentObjectId = _
        VTPowerPointShapeObjectIdentity(currentSlide, selectedShape)
    On Error Resume Next
    ownerDocumentId = selectedShape.Tags(VT_POWERPOINT_DOCUMENT_TAG)
    ownerObjectId = selectedShape.Tags(VT_POWERPOINT_OBJECT_TAG)
    On Error GoTo 0

    ' Current formulas carry both the presentation identity and the physical
    ' SlideID:Shape.ID that owned the formula when it was committed. Copy/paste
    ' duplicates tags verbatim but PowerPoint gives the pasted Shape a new ID,
    ' so only the physical copy fails this comparison.
    If Len(ownerDocumentId) > 0 And ownerDocumentId <> currentDocumentId Then
        VTShouldForkPowerPointFormulaCopy = True
        Exit Function
    End If
    If Len(ownerObjectId) > 0 Then
        VTShouldForkPowerPointFormulaCopy = _
            (ownerDocumentId <> currentDocumentId) Or _
            (ownerObjectId <> currentObjectId)
        Exit Function
    End If

    ' Lazily migrate formulas created before the physical-owner tag existed.
    ' Keep one deterministic original (first slide/order occurrence) and stamp
    ' that Shape only. All other physical copies fork when edited.
    For Each slideObject In ActivePresentation.Slides
        For Each candidate In slideObject.Shapes
            If VTShapeFormulaId(candidate) = formulaId Then
                matchCount = matchCount + 1
                If canonicalShape Is Nothing Then
                    Set canonicalSlide = slideObject
                    Set canonicalShape = candidate
                End If
            End If
        Next candidate
    Next slideObject
    If canonicalShape Is Nothing Then
        VTShouldForkPowerPointFormulaCopy = True
        Exit Function
    End If
    VTSetShapeTag canonicalShape, VT_POWERPOINT_DOCUMENT_TAG, currentDocumentId
    VTSetShapeTag canonicalShape, VT_POWERPOINT_OBJECT_TAG, _
        VTPowerPointShapeObjectIdentity(canonicalSlide, canonicalShape)
    VTShouldForkPowerPointFormulaCopy = _
        VTPowerPointShapeObjectIdentity(canonicalSlide, canonicalShape) <> _
        currentObjectId
End Function

Private Sub VTPowerPointEditShape(ByVal selectedShape As Shape)
    Dim formulaId As String
    Dim encodedMetadata As String
    Dim formulaReference As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim sessionId As String
    Dim requestJson As String
    Dim powerPointJson As String
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim launchTiming As String
    Dim startedAt As Double
    Dim envelopeReadyAt As Double
    Dim scaleReadyAt As Double
    Dim identifierReadyAt As Double
    Dim geometryReadyAt As Double
    Dim requestReadyAt As Double
    Dim launchFinishedAt As Double
    Dim forkCopiedFormula As Boolean

    startedAt = Timer
    If selectedShape Is Nothing Then
        Err.Raise vbObjectError + 7500, "VisualTeX", "Select one VisualTeX formula shape."
    End If
    formulaId = VTShapeFormulaId(selectedShape)
    encodedMetadata = VTShapeMetadata(selectedShape)
    formulaReference = selectedShape.Title
    VTValidateEditEnvelope encodedMetadata, formulaReference, formulaId, displayMode, numbered
    If Len(formulaId) = 0 Then formulaId = VTShapeFormulaId(selectedShape)
    If Len(formulaId) = 0 Then
        Err.Raise vbObjectError + 7500, "VisualTeX", "The selected PowerPoint shape has no VisualTeX formula id."
    End If
    envelopeReadyAt = Timer
    VTEnsurePowerPointFormulaScaleState _
        selectedShape, fontSizePt, referenceWidthPt, referenceHeightPt
    scaleReadyAt = Timer

    forkCopiedFormula = _
        VTShouldForkPowerPointFormulaCopy( _
            ActiveWindow.View.Slide, selectedShape, formulaId)
    If forkCopiedFormula Then
        formulaId = VTNewUuidV4()
        selectedShape.Name = VT_SHAPE_PREFIX & formulaId
    End If
    sessionId = VTNewUuidV4()
    identifierReadyAt = Timer
    powerPointJson = VTPowerPointGeometryJson( _
        ActiveWindow.View.Slide, selectedShape, fontSizePt, _
        referenceWidthPt, referenceHeightPt)
    geometryReadyAt = Timer
    requestJson = VTRequestJson( _
        sessionId, _
        VT_POWERPOINT_HOST, _
        "edit", _
        formulaId, _
        "block", _
        False, _
        VTPresentationIdentity(), _
        selectedShape.Name, _
        encodedMetadata, _
        "", _
        powerPointJson, _
        False, _
        0#, _
        0#, _
        0#, _
        "formula", _
        forkCopiedFormula)
    requestReadyAt = Timer
    launchTiming = VTWriteAndLaunchSession( _
        VT_POWERPOINT_HOST, sessionId, requestJson)
    launchFinishedAt = Timer
    VTWritePowerPointOpenPerformance _
        sessionId, startedAt, envelopeReadyAt, scaleReadyAt, _
        identifierReadyAt, geometryReadyAt, requestReadyAt, _
        launchFinishedAt, launchTiming
End Sub

Private Function VTTimerElapsedMilliseconds( _
    ByVal startedAt As Double, _
    ByVal finishedAt As Double) As Double

    If finishedAt < startedAt Then finishedAt = finishedAt + 86400#
    VTTimerElapsedMilliseconds = (finishedAt - startedAt) * 1000#
End Function

Private Sub VTWritePowerPointOpenPerformance( _
    ByVal sessionId As String, _
    ByVal startedAt As Double, _
    ByVal envelopeReadyAt As Double, _
    ByVal scaleReadyAt As Double, _
    ByVal identifierReadyAt As Double, _
    ByVal geometryReadyAt As Double, _
    ByVal requestReadyAt As Double, _
    ByVal launchFinishedAt As Double, _
    ByVal launchTiming As String)

    Dim tracePath As String
    Dim payload As String

    On Error GoTo TraceFinished
    tracePath = VTSessionDirectory(sessionId) & _
        "/vba-open-performance.txt"
    payload = _
        "envelopeMs=" & VTJsonNumber( _
            VTTimerElapsedMilliseconds(startedAt, envelopeReadyAt)) & ";" & _
        "scaleMs=" & VTJsonNumber( _
            VTTimerElapsedMilliseconds(envelopeReadyAt, scaleReadyAt)) & ";" & _
        "identifierMs=" & VTJsonNumber( _
            VTTimerElapsedMilliseconds(scaleReadyAt, identifierReadyAt)) & ";" & _
        "geometryMs=" & VTJsonNumber( _
            VTTimerElapsedMilliseconds(identifierReadyAt, geometryReadyAt)) & ";" & _
        "requestMs=" & VTJsonNumber( _
            VTTimerElapsedMilliseconds(geometryReadyAt, requestReadyAt)) & ";" & _
        "launcherCallMs=" & VTJsonNumber( _
            VTTimerElapsedMilliseconds(requestReadyAt, launchFinishedAt)) & ";" & _
        "totalMs=" & VTJsonNumber( _
            VTTimerElapsedMilliseconds(startedAt, launchFinishedAt)) & ";" & _
        launchTiming
    VTWriteTextAtomic tracePath, payload

TraceFinished:
    Err.Clear
End Sub

Public Sub VisualTeX_PreparePowerPointMovedCopyRegression()
    Dim sourceShape As Shape
    Dim copiedRange As ShapeRange
    Dim copiedShape As Shape
    Dim resultPath As String

    VTRequireWritablePowerPointPresentation
    Set sourceShape = VTSelectedSingleShape()
    If Not VTIsVisualTeXPowerPointShape(sourceShape) Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "Select one committed VisualTeX PowerPoint formula."
    End If

    VT_MOVED_COPY_TEST_SOURCE_ID = sourceShape.Id
    VT_MOVED_COPY_TEST_SOURCE_LEFT = sourceShape.Left
    VT_MOVED_COPY_TEST_SOURCE_TOP = sourceShape.Top
    VT_MOVED_COPY_TEST_SOURCE_FORMULA_ID = VTShapeFormulaId(sourceShape)

    Set copiedRange = sourceShape.Duplicate
    If copiedRange.Count <> 1 Then
        Err.Raise vbObjectError + 7606, "VisualTeX regression", _
            "PowerPoint did not create exactly one moved-copy test Shape."
    End If
    Set copiedShape = copiedRange(1)
    copiedShape.Left = sourceShape.Left + 210!
    copiedShape.Top = sourceShape.Top + 120!
    VT_MOVED_COPY_TEST_COPIED_ID = copiedShape.Id
    VT_MOVED_COPY_TEST_COPIED_LEFT = copiedShape.Left
    VT_MOVED_COPY_TEST_COPIED_TOP = copiedShape.Top
    copiedShape.Select

    resultPath = VTApplicationSupportRoot() & _
        "/Tests/powerpoint-moved-copy-regression-result.txt"
    VTWriteTextAtomic _
        resultPath, _
        "PREPARED" & vbLf & _
        "sourceId=" & CStr(VT_MOVED_COPY_TEST_SOURCE_ID) & vbLf & _
        "copiedId=" & CStr(VT_MOVED_COPY_TEST_COPIED_ID) & vbLf & _
        "sourceName=" & sourceShape.Name & vbLf & _
        "copiedName=" & copiedShape.Name & vbLf & _
        "sourceLeft=" & VTJsonNumber(VT_MOVED_COPY_TEST_SOURCE_LEFT) & vbLf & _
        "sourceTop=" & VTJsonNumber(VT_MOVED_COPY_TEST_SOURCE_TOP) & vbLf & _
        "copiedLeft=" & VTJsonNumber(VT_MOVED_COPY_TEST_COPIED_LEFT) & vbLf & _
        "copiedTop=" & VTJsonNumber(VT_MOVED_COPY_TEST_COPIED_TOP) & vbLf
End Sub

Public Sub VisualTeX_VerifyPowerPointMovedCopyRegression()
    Const positionTolerance As Double = 0.1

    Dim currentSlide As Slide
    Dim sourceShape As Shape
    Dim candidate As Shape
    Dim movedCopy As Shape
    Dim resultPath As String

    VTRequireWritablePowerPointPresentation
    Set currentSlide = ActiveWindow.View.Slide
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/powerpoint-moved-copy-regression-result.txt"

    Set sourceShape = _
        VTFindPowerPointShapeById(currentSlide, VT_MOVED_COPY_TEST_SOURCE_ID)
    If sourceShape Is Nothing Then
        Err.Raise vbObjectError + 7607, "VisualTeX regression", _
            "The source PowerPoint formula disappeared after editing its copy."
    End If
    If Abs(sourceShape.Left - VT_MOVED_COPY_TEST_SOURCE_LEFT) > _
            positionTolerance Or _
       Abs(sourceShape.Top - VT_MOVED_COPY_TEST_SOURCE_TOP) > _
            positionTolerance Then
        Err.Raise vbObjectError + 7607, "VisualTeX regression", _
            "Editing the moved copy changed the source formula position."
    End If

    For Each candidate In currentSlide.Shapes
        If candidate.Id <> sourceShape.Id Then
            If Abs(candidate.Left - VT_MOVED_COPY_TEST_COPIED_LEFT) <= _
                    positionTolerance And _
               Abs(candidate.Top - VT_MOVED_COPY_TEST_COPIED_TOP) <= _
                    positionTolerance Then
                If movedCopy Is Nothing Then
                    Set movedCopy = candidate
                Else
                    Err.Raise vbObjectError + 7607, "VisualTeX regression", _
                        "More than one Shape matched the moved-copy position."
                End If
            End If
        End If
    Next candidate
    If movedCopy Is Nothing Then
        Err.Raise vbObjectError + 7607, "VisualTeX regression", _
            "The edited copy did not remain at its moved position."
    End If

    VTWriteTextAtomic _
        resultPath, _
        "PASS" & vbLf & _
        "sourceId=" & CStr(sourceShape.Id) & vbLf & _
        "movedCopyId=" & CStr(movedCopy.Id) & vbLf & _
        "sourceLeft=" & VTJsonNumber(sourceShape.Left) & vbLf & _
        "sourceTop=" & VTJsonNumber(sourceShape.Top) & vbLf & _
        "movedCopyLeft=" & VTJsonNumber(movedCopy.Left) & vbLf & _
        "movedCopyTop=" & VTJsonNumber(movedCopy.Top) & vbLf
End Sub

Public Sub VisualTeX_RunPowerPointCopyIdentityRegression()
    Const formulaId As String = _
        "68686868-6868-4868-8868-686868686868"

    Dim sourcePresentation As Presentation
    Dim testPresentation As Presentation
    Dim testSlide As Slide
    Dim sourceShape As Shape
    Dim copiedShape As Shape
    Dim copiedRange As ShapeRange
    Dim encodedMetadata As String
    Dim formulaReference As String
    Dim sourceObjectId As String
    Dim copiedStoredObjectId As String
    Dim copiedActualObjectId As String
    Dim resultPath As String
    Dim regressionStage As String
    Dim regressionErrorNumber As Long
    Dim regressionErrorDescription As String

    On Error GoTo RegressionFailed
    If Presentations.Count > 0 Then Set sourcePresentation = ActivePresentation
    resultPath = VTApplicationSupportRoot() & _
        "/Tests/powerpoint-copy-identity-regression-result.txt"
    encodedMetadata = VT_METADATA_PREFIX & "e30"
    formulaReference = VTFormulaReference(formulaId, "block", False)

    regressionStage = "create-source"
    Set testPresentation = Presentations.Add(msoTrue)
    Set testSlide = testPresentation.Slides.Add(1, ppLayoutBlank)
    Set sourceShape = testSlide.Shapes.AddShape( _
        msoShapeRectangle, 72!, 72!, 160!, 48!)
    sourceShape.Name = VT_SHAPE_PREFIX & formulaId
    sourceShape.Title = formulaReference
    sourceShape.AlternativeText = encodedMetadata
    VTSetShapeTag sourceShape, "VisualTeXFormulaId", formulaId
    VTSetShapeTag sourceShape, VT_POWERPOINT_DOCUMENT_TAG, _
        VTPresentationIdentity()
    sourceObjectId = VTPowerPointShapeObjectIdentity(testSlide, sourceShape)
    VTSetShapeTag sourceShape, VT_POWERPOINT_OBJECT_TAG, sourceObjectId

    regressionStage = "duplicate"
    Set copiedRange = sourceShape.Duplicate
    If copiedRange.Count <> 1 Then
        Err.Raise vbObjectError + 7605, "VisualTeX regression", _
            "PowerPoint did not create exactly one duplicate Shape."
    End If
    Set copiedShape = copiedRange(1)
    copiedStoredObjectId = copiedShape.Tags(VT_POWERPOINT_OBJECT_TAG)
    copiedActualObjectId = _
        VTPowerPointShapeObjectIdentity(testSlide, copiedShape)
    If sourceShape.Id = copiedShape.Id Or _
       copiedStoredObjectId <> sourceObjectId Or _
       copiedActualObjectId = sourceObjectId Then
        Err.Raise vbObjectError + 7605, "VisualTeX regression", _
            "PowerPoint did not preserve copied tags with a distinct physical Shape.ID."
    End If

    regressionStage = "classify"
    If VTShouldForkPowerPointFormulaCopy( _
       testSlide, sourceShape, formulaId) Then
        Err.Raise vbObjectError + 7605, "VisualTeX regression", _
            "The original PowerPoint formula was incorrectly classified as a copy."
    End If
    If Not VTShouldForkPowerPointFormulaCopy( _
       testSlide, copiedShape, formulaId) Then
        Err.Raise vbObjectError + 7605, "VisualTeX regression", _
            "The copied PowerPoint formula retained the source physical identity."
    End If

    VTWriteTextAtomic _
        resultPath, _
        "PASS" & vbLf & _
        "revision=" & VT_POWERPOINT_SOURCE_REVISION & vbLf & _
        "sourceObjectId=" & sourceObjectId & vbLf & _
        "copiedStoredObjectId=" & copiedStoredObjectId & vbLf & _
        "copiedActualObjectId=" & copiedActualObjectId & vbLf
    testPresentation.Close
    If Not sourcePresentation Is Nothing Then sourcePresentation.Windows(1).Activate
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
    If Not testPresentation Is Nothing Then testPresentation.Close
    If Not sourcePresentation Is Nothing Then sourcePresentation.Windows(1).Activate
    On Error GoTo 0
    Err.Raise regressionErrorNumber, _
        "VisualTeX PowerPoint copy identity regression", _
        regressionStage & ": " & regressionErrorDescription
End Sub

Public Sub VisualTeX_DeleteSelected()
    On Error GoTo Failed
    Dim selectedShape As Shape
    VTRequireWritablePowerPointPresentation
    Set selectedShape = VTSelectedSingleShape()
    If Len(VTShapeFormulaId(selectedShape)) = 0 Then
        Err.Raise vbObjectError + 7501, "VisualTeX", "Select one VisualTeX formula shape."
    End If
    selectedShape.Delete
    Exit Sub
Failed:
    VTShowError "PowerPoint delete", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_OpenApplication()
    On Error GoTo Failed
    VTOpenApplication VT_POWERPOINT_HOST
    Exit Sub
Failed:
    VTShowError "application launch", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_ApplyPendingResult()
    On Error GoTo Failed

    Dim sessionId As String
    Dim dispatch As Object
    Dim actionName As String
    Dim hostName As String

    sessionId = VTReadActiveSessionId(VT_POWERPOINT_HOST)
    Set dispatch = VTReadDispatch(sessionId)
    actionName = CStr(dispatch("action"))
    hostName = CStr(dispatch("host"))
    If hostName <> VT_POWERPOINT_HOST Then
        Err.Raise vbObjectError + 7502, "VisualTeX", "The active VisualTeX dispatch is not for PowerPoint."
    End If

    Select Case actionName
        Case "commit": VTFinalizePowerPointDispatch sessionId, dispatch
        Case "cancel": VTCancelPowerPointDispatch sessionId, dispatch
        Case Else
            Err.Raise vbObjectError + 7503, "VisualTeX", "The VisualTeX PowerPoint dispatch action is invalid."
    End Select
    Exit Sub

Failed:
    Err.Raise Err.Number, "VisualTeX PowerPoint callback", Err.Description
End Sub

Private Sub VTFinalizePowerPointDispatch(ByVal sessionId As String, ByVal dispatch As Object)
    Dim formulaId As String
    Dim metadata As String
    Dim shapeName As String
    Dim sourceShapeIndex As Long
    Dim sourceShapeId As Long
    Dim sourceShapeName As String
    Dim imagePath As String
    Dim fallbackImagePath As String
    Dim vectorInsertErrorNumber As Long
    Dim vectorInsertErrorDescription As String
    Dim expectedPresentation As String
    Dim slideIndex As Long
    Dim slideId As Long
    Dim targetZOrder As Long
    Dim targetLeft As Double
    Dim targetTop As Double
    Dim targetWidth As Double
    Dim targetHeight As Double
    Dim targetRotation As Double
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim targetPresentation As Presentation
    Dim currentSlide As Slide
    Dim committed As Shape
    Dim original As Shape
    Dim candidate As Shape
    Dim originalTemporaryName As String
    Dim candidateTemporaryName As String
    Dim originalRenamed As Boolean
    Dim formulaReference As String
    Dim existingSessionId As String
    Dim existingPendingState As String

    VTRequireDispatchValue dispatch, "formulaId"
    VTRequireDispatchValue dispatch, "metadata"
    VTRequireDispatchValue dispatch, "shapeName"
    VTRequireDispatchValue dispatch, "sourceShapeName"
    VTRequireDispatchValue dispatch, "imagePath"
    VTRequireDispatchValue dispatch, "presentationIdentity"
    VTRequireDispatchValue dispatch, "slideIndex"
    VTRequireDispatchValue dispatch, "slideId"
    VTRequireDispatchValue dispatch, "targetLeft"
    VTRequireDispatchValue dispatch, "targetTop"
    VTRequireDispatchValue dispatch, "targetWidth"
    VTRequireDispatchValue dispatch, "targetHeight"
    VTRequireDispatchValue dispatch, "fontSizePt"
    VTRequireDispatchValue dispatch, "referenceWidthPt"
    VTRequireDispatchValue dispatch, "referenceHeightPt"
    VTRequireDispatchValue dispatch, "rotation"
    VTRequireDispatchValue dispatch, "zOrder"

    formulaId = CStr(dispatch("formulaId"))
    metadata = CStr(dispatch("metadata"))
    shapeName = CStr(dispatch("shapeName"))
    sourceShapeName = CStr(dispatch("sourceShapeName"))
    imagePath = CStr(dispatch("imagePath"))
    fallbackImagePath = VTDispatchOptionalPpt(dispatch, "fallbackImagePath")
    expectedPresentation = CStr(dispatch("presentationIdentity"))
    Set targetPresentation = VTFindPowerPointPresentation(expectedPresentation)
    VTRequireWritablePowerPointPresentationObject targetPresentation
    slideIndex = CLng(dispatch("slideIndex"))
    slideId = CLng(dispatch("slideId"))
    targetLeft = VTDispatchDoublePpt(dispatch, "targetLeft")
    targetTop = VTDispatchDoublePpt(dispatch, "targetTop")
    targetWidth = VTDispatchPositiveDoublePpt(dispatch, "targetWidth")
    targetHeight = VTDispatchPositiveDoublePpt(dispatch, "targetHeight")
    fontSizePt = VTDispatchPositiveDoublePpt(dispatch, "fontSizePt")
    referenceWidthPt = VTDispatchPositiveDoublePpt( _
        dispatch, "referenceWidthPt")
    referenceHeightPt = VTDispatchPositiveDoublePpt( _
        dispatch, "referenceHeightPt")
    targetRotation = VTDispatchDoublePpt(dispatch, "rotation")
    targetZOrder = CLng(dispatch("zOrder"))

    If Not VTIsCanonicalUuid(formulaId) Or _
       shapeName <> VT_SHAPE_PREFIX & formulaId Or _
       Not VTIsEncodedMetadata(metadata) Or _
       Not VTValidPowerPointFormulaFontSize(fontSizePt) Or _
       referenceWidthPt <= 0# Or referenceHeightPt <= 0# Then
        Err.Raise vbObjectError + 7504, "VisualTeX", _
            "VisualTeX PowerPoint result metadata is invalid."
    End If
    formulaReference = VTFormulaReference(formulaId, "block", False)
    If slideIndex <= 0 Or slideIndex > targetPresentation.Slides.Count Or slideId <= 0 Or targetZOrder <= 0 Then
        Err.Raise vbObjectError + 7516, "VisualTeX", "VisualTeX PowerPoint target reference is invalid."
    End If
    VTValidateAbsoluteVisualTeXPath imagePath
    If Not VTPathFileExists(imagePath) Then
        Err.Raise vbObjectError + 7517, "VisualTeX", "VisualTeX PowerPoint SVG result is missing."
    End If
    If Len(fallbackImagePath) > 0 Then
        VTValidateAbsoluteVisualTeXPath fallbackImagePath
        If Not VTPathFileExists(fallbackImagePath) Then fallbackImagePath = ""
    End If

    Set currentSlide = targetPresentation.Slides(slideIndex)
    If currentSlide.SlideID <> slideId Then
        Err.Raise vbObjectError + 7518, "VisualTeX", "The original PowerPoint slide no longer exists."
    End If
    On Error Resume Next
    Set committed = currentSlide.Shapes(shapeName)
    If Not committed Is Nothing Then
        existingSessionId = committed.Tags("VisualTeXSessionId")
        existingPendingState = committed.Tags("VisualTeXPending")
    End If
    Err.Clear
    On Error GoTo TransactionFailed
    If Not committed Is Nothing Then
        ' A normal edit starts from the previously committed shape, whose
        ' Session id differs from this Apply. Avoid reading every geometry and
        ' metadata property unless this is an actual retry of the same commit.
        If existingSessionId = sessionId And existingPendingState = "0" Then
            If VTIsCommittedPowerPointShape( _
                committed, shapeName, formulaReference, metadata, formulaId, _
                sessionId, targetLeft, targetTop, targetWidth, targetHeight, _
                targetRotation, targetZOrder, fontSizePt, referenceWidthPt, _
                referenceHeightPt) Then
                Exit Sub
            End If
        End If
    End If
    Set committed = Nothing

    On Error Resume Next
    Set original = currentSlide.Shapes(sourceShapeName)
    On Error GoTo TransactionFailed
    If original Is Nothing Then
        Err.Raise vbObjectError + 7519, "VisualTeX", _
            "The original VisualTeX PowerPoint shape no longer exists."
    End If

    candidateTemporaryName = "VisualTeXPendingResult_" & Replace$(Left$(sessionId, 13), "-", "")
    originalTemporaryName = "VisualTeXOriginal_" & Replace$(Left$(sessionId, 13), "-", "")
    ' Modern PowerPoint for Mac preserves an imported SVG as vector artwork.
    ' Prefer SVG so formulas remain sharp at arbitrary zoom. Keep the PNG only
    ' as a compatibility fallback for Office builds that reject SVG AddPicture.
    DoEvents
    On Error Resume Next
    Set candidate = currentSlide.Shapes.AddPicture( _
        FileName:=imagePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=CSng(targetLeft), _
        Top:=CSng(targetTop), _
        Width:=CSng(targetWidth), _
        Height:=CSng(targetHeight))
    vectorInsertErrorNumber = Err.Number
    vectorInsertErrorDescription = Err.Description
    Err.Clear
    DoEvents
    On Error GoTo TransactionFailed
    If candidate Is Nothing Then
        If Len(fallbackImagePath) = 0 Then
            Err.Raise vbObjectError + 7521, "VisualTeX", _
                "PowerPoint could not insert the VisualTeX SVG: " & _
                CStr(vectorInsertErrorNumber) & " " & vectorInsertErrorDescription
        End If
        DoEvents
        Set candidate = currentSlide.Shapes.AddPicture( _
            FileName:=fallbackImagePath, _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=CSng(targetLeft), _
            Top:=CSng(targetTop), _
            Width:=CSng(targetWidth), _
            Height:=CSng(targetHeight))
        DoEvents
    End If
    candidate.Name = candidateTemporaryName
    ' AddPicture already received the final bounds. Reassigning all four
    ' geometry properties forces redundant Office layout work on every Apply.
    candidate.LockAspectRatio = msoTrue
    If Abs(targetRotation) > 0.01 Then
        candidate.Rotation = CSng(targetRotation)
    End If
    candidate.AlternativeText = metadata
    candidate.Title = formulaReference
    VTSetShapeTag candidate, "VisualTeXFormulaId", formulaId
    VTSetShapeTag candidate, VT_POWERPOINT_DOCUMENT_TAG, expectedPresentation
    VTSetShapeTag candidate, VT_POWERPOINT_OBJECT_TAG, _
        VTPowerPointShapeObjectIdentity(currentSlide, candidate)
    VTSetShapeTag candidate, "VisualTeXSessionId", sessionId
    VTSetShapeTag candidate, "VisualTeXPending", "0"
    VTSetPowerPointFormulaScaleState _
        candidate, fontSizePt, referenceWidthPt, referenceHeightPt
    On Error Resume Next
    VTSetShapeTag candidate, "VisualTeXMetadata", metadata
    Err.Clear
    On Error GoTo TransactionFailed
    DoEvents

    ' The original still occupies targetZOrder. Put the candidate immediately
    ' above it; deleting the original as the final mutation shifts the candidate
    ' into the exact original z-order without any fallible operation afterwards.
    VTRestoreZOrder candidate, targetZOrder + 1
    DoEvents
    original.Name = originalTemporaryName
    originalRenamed = True
    candidate.Name = shapeName
    DoEvents

    ' Keep the synchronous transaction check focused on identity and rollback
    ' safety. The integration benchmark verifies geometry and scale metadata
    ' after the callback without extending the user's Apply wait.
    If candidate.Name <> shapeName Or _
       candidate.ZOrderPosition <> targetZOrder + 1 Or _
       candidate.AlternativeText <> metadata Or _
       candidate.Title <> formulaReference Or _
       candidate.Tags("VisualTeXFormulaId") <> formulaId Or _
       candidate.Tags("VisualTeXSessionId") <> sessionId Or _
       candidate.Tags("VisualTeXPending") <> "0" Then
        Err.Raise vbObjectError + 7520, "VisualTeX", _
            "PowerPoint did not persist the VisualTeX formula identity."
    End If

    DoEvents
    original.Delete
    Exit Sub

TransactionFailed:
    Dim transactionErrorNumber As Long
    Dim transactionErrorDescription As String
    transactionErrorNumber = Err.Number
    transactionErrorDescription = Err.Description
    On Error Resume Next
    If Not candidate Is Nothing Then candidate.Delete
    If originalRenamed And Not original Is Nothing Then original.Name = sourceShapeName
    On Error GoTo 0
    Err.Raise transactionErrorNumber, "VisualTeX PowerPoint transaction", transactionErrorDescription
End Sub

Private Function VTIsCommittedPowerPointShape( _
    ByVal target As Shape, _
    ByVal expectedName As String, _
    ByVal formulaReference As String, _
    ByVal metadata As String, _
    ByVal formulaId As String, _
    ByVal sessionId As String, _
    ByVal targetLeft As Double, _
    ByVal targetTop As Double, _
    ByVal targetWidth As Double, _
    ByVal targetHeight As Double, _
    ByVal targetRotation As Double, _
    ByVal targetZOrder As Long, _
    ByVal fontSizePt As Double, _
    ByVal referenceWidthPt As Double, _
    ByVal referenceHeightPt As Double) As Boolean

    On Error GoTo NotCommitted
    ' VBA does not short-circuit And expressions. Check cheap identifiers first
    ' so a retry mismatch does not read every COM-backed shape property.
    If target.Name <> expectedName Then Exit Function
    If target.Tags("VisualTeXSessionId") <> sessionId Then Exit Function
    If target.Tags("VisualTeXPending") <> "0" Then Exit Function
    If target.Tags("VisualTeXFormulaId") <> formulaId Then Exit Function
    If target.AlternativeText <> metadata Then Exit Function
    If target.Title <> formulaReference Then Exit Function
    If target.ZOrderPosition <> targetZOrder Then Exit Function
    If Abs(target.Left - targetLeft) > 0.1 Then Exit Function
    If Abs(target.Top - targetTop) > 0.1 Then Exit Function
    If Abs(target.Width - targetWidth) > 0.1 Then Exit Function
    If Abs(target.Height - targetHeight) > 0.1 Then Exit Function
    If Abs(target.Rotation - targetRotation) > 0.1 Then Exit Function
    If target.Tags(VT_POWERPOINT_FONT_TAG) <> _
       VTJsonNumber(fontSizePt) Then Exit Function
    If target.Tags(VT_POWERPOINT_REFERENCE_WIDTH_TAG) <> _
       VTJsonNumber(referenceWidthPt) Then Exit Function
    If target.Tags(VT_POWERPOINT_REFERENCE_HEIGHT_TAG) <> _
       VTJsonNumber(referenceHeightPt) Then Exit Function
    VTIsCommittedPowerPointShape = True
    Exit Function

NotCommitted:
    Err.Clear
    VTIsCommittedPowerPointShape = False
End Function

Private Sub VTCancelPowerPointDispatch(ByVal sessionId As String, ByVal dispatch As Object)
    Dim pendingMarker As String
    Dim currentPresentation As Presentation
    Dim currentSlide As Slide
    Dim shapeItem As Shape

    pendingMarker = VTDispatchOptionalPpt(dispatch, "pendingMarker")
    If Len(pendingMarker) = 0 Or Presentations.Count = 0 Then Exit Sub
    On Error Resume Next
    For Each currentPresentation In Presentations
        For Each currentSlide In currentPresentation.Slides
            For Each shapeItem In currentSlide.Shapes
                If shapeItem.AlternativeText = pendingMarker And _
                   shapeItem.Tags("VisualTeXSessionId") = sessionId And _
                   shapeItem.Tags("VisualTeXPending") = "1" Then
                    shapeItem.Delete
                    Exit Sub
                End If
            Next shapeItem
        Next currentSlide
    Next currentPresentation
    On Error GoTo 0
End Sub

Private Function VTSelectedSingleShape() As Shape
    If ActiveWindow Is Nothing Then
        Err.Raise vbObjectError + 7505, "VisualTeX", "PowerPoint has no active window."
    End If
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        Err.Raise vbObjectError + 7506, "VisualTeX", "Select exactly one VisualTeX formula shape."
    End If
    If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
        Err.Raise vbObjectError + 7507, "VisualTeX", "Select exactly one VisualTeX formula shape."
    End If
    Set VTSelectedSingleShape = ActiveWindow.Selection.ShapeRange(1)
End Function

Private Function VTShapeFormulaId(ByVal target As Shape) As String
    Dim value As String
    On Error Resume Next
    value = target.Tags("VisualTeXFormulaId")
    On Error GoTo 0
    If Not VTIsCanonicalUuid(value) And Left$(target.Name, Len(VT_SHAPE_PREFIX)) = VT_SHAPE_PREFIX Then
        value = Mid$(target.Name, Len(VT_SHAPE_PREFIX) + 1)
    End If
    If VTIsCanonicalUuid(value) Then VTShapeFormulaId = value
End Function

Private Function VTShapeMetadata(ByVal target As Shape) As String
    Dim value As String
    On Error Resume Next
    value = target.Tags("VisualTeXMetadata")
    On Error GoTo 0
    If Not VTIsEncodedMetadata(value) Then value = target.AlternativeText
    If VTIsEncodedMetadata(value) Then VTShapeMetadata = value
End Function

Private Function VTFindUniqueFormulaShape(ByVal shapeName As String) As Shape
    Dim slideItem As Slide
    Dim candidate As Shape
    Dim match As Shape
    Dim count As Long

    If Len(shapeName) = 0 Or Len(shapeName) > 128 Or InStr(shapeName, vbCr) > 0 Or InStr(shapeName, vbLf) > 0 Then
        Err.Raise vbObjectError + 7508, "VisualTeX", "VisualTeX PowerPoint shape name is invalid."
    End If
    For Each slideItem In ActivePresentation.Slides
        Set candidate = Nothing
        On Error Resume Next
        Set candidate = slideItem.Shapes(shapeName)
        Err.Clear
        On Error GoTo 0
        If Not candidate Is Nothing Then
            count = count + 1
            Set match = candidate
        End If
    Next slideItem
    If count <> 1 Or match Is Nothing Then
        Err.Raise vbObjectError + 7509, "VisualTeX", "PowerPoint must contain exactly one matching VisualTeX formula shape."
    End If
    Set VTFindUniqueFormulaShape = match
End Function

Private Function VTValidPowerPointFormulaFontSize( _
    ByVal value As Double) As Boolean

    VTValidPowerPointFormulaFontSize = _
        value >= VT_POWERPOINT_MIN_FONT_SIZE_PT And _
        value <= VT_POWERPOINT_MAX_FONT_SIZE_PT
End Function

Private Function VTTryReadPowerPointShapeTagDouble( _
    ByVal target As Shape, _
    ByVal tagName As String, _
    ByRef value As Double) As Boolean

    Dim storedValue As String

    value = 0#
    If target Is Nothing Or Len(tagName) = 0 Then Exit Function
    On Error GoTo InvalidValue
    storedValue = target.Tags(tagName)
    If Len(storedValue) = 0 Then Exit Function
    value = VTParseInvariantDouble(storedValue)
    If Abs(value) > 10000000# Then Exit Function
    VTTryReadPowerPointShapeTagDouble = True
    Exit Function

InvalidValue:
    value = 0#
    Err.Clear
End Function

Private Sub VTSetPowerPointFormulaScaleState( _
    ByVal target As Shape, _
    ByVal fontSizePt As Double, _
    ByVal referenceWidthPt As Double, _
    ByVal referenceHeightPt As Double)

    If target Is Nothing Or _
       Not VTValidPowerPointFormulaFontSize(fontSizePt) Or _
       referenceWidthPt <= 0# Or referenceHeightPt <= 0# Or _
       referenceWidthPt > 10000# Or referenceHeightPt > 10000# Then
        Err.Raise vbObjectError + 7523, "VisualTeX", _
            "The PowerPoint formula point-size state is invalid."
    End If
    VTSetShapeTag target, VT_POWERPOINT_FONT_TAG, VTJsonNumber(fontSizePt)
    VTSetShapeTag target, VT_POWERPOINT_REFERENCE_WIDTH_TAG, _
        VTJsonNumber(referenceWidthPt)
    VTSetShapeTag target, VT_POWERPOINT_REFERENCE_HEIGHT_TAG, _
        VTJsonNumber(referenceHeightPt)
End Sub

Private Sub VTEnsurePowerPointFormulaScaleState( _
    ByVal target As Shape, _
    ByRef fontSizePt As Double, _
    ByRef referenceWidthPt As Double, _
    ByRef referenceHeightPt As Double)

    Dim hasFontSize As Boolean
    Dim hasReferenceWidth As Boolean
    Dim hasReferenceHeight As Boolean
    Dim observedFontSizePt As Double
    Dim scaleStateNeedsWrite As Boolean

    If target Is Nothing Or Not VTIsVisualTeXPowerPointShape(target) Then
        Err.Raise vbObjectError + 7524, "VisualTeX", _
            "Select one VisualTeX PowerPoint SVG formula."
    End If
    hasFontSize = VTTryReadPowerPointShapeTagDouble( _
        target, VT_POWERPOINT_FONT_TAG, fontSizePt)
    hasReferenceWidth = VTTryReadPowerPointShapeTagDouble( _
        target, VT_POWERPOINT_REFERENCE_WIDTH_TAG, referenceWidthPt)
    hasReferenceHeight = VTTryReadPowerPointShapeTagDouble( _
        target, VT_POWERPOINT_REFERENCE_HEIGHT_TAG, referenceHeightPt)

    If Not hasFontSize Or _
       Not VTValidPowerPointFormulaFontSize(fontSizePt) Then
        fontSizePt = VT_POWERPOINT_DEFAULT_FONT_SIZE_PT
        scaleStateNeedsWrite = True
    End If
    If Not hasReferenceWidth Or Not hasReferenceHeight Or _
       referenceWidthPt <= 0# Or referenceHeightPt <= 0# Then
        referenceWidthPt = target.Width * _
            VT_POWERPOINT_REFERENCE_FONT_SIZE_PT / fontSizePt
        referenceHeightPt = target.Height * _
            VT_POWERPOINT_REFERENCE_FONT_SIZE_PT / fontSizePt
        scaleStateNeedsWrite = True
    Else
        observedFontSizePt = target.Height / referenceHeightPt * _
            VT_POWERPOINT_REFERENCE_FONT_SIZE_PT
        If VTValidPowerPointFormulaFontSize(observedFontSizePt) Then
            If Abs(observedFontSizePt - fontSizePt) > 0.05 Then
                fontSizePt = observedFontSizePt
                scaleStateNeedsWrite = True
            End If
        End If
    End If
    If scaleStateNeedsWrite Then
        VTSetPowerPointFormulaScaleState _
            target, fontSizePt, referenceWidthPt, referenceHeightPt
    End If
End Sub

Private Function VTPreferredPowerPointFormulaFontSize() As Double
    Dim selectedShape As Shape
    Dim selectedSize As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double

    selectedSize = VT_POWERPOINT_DEFAULT_FONT_SIZE_PT
    On Error GoTo UseDefault
    If ActiveWindow Is Nothing Then GoTo UseDefault
    If ActiveWindow.Selection.Type = ppSelectionText Then
        selectedSize = ActiveWindow.Selection.TextRange.Font.Size
    ElseIf ActiveWindow.Selection.Type = ppSelectionShapes And _
           ActiveWindow.Selection.ShapeRange.Count = 1 Then
        Set selectedShape = ActiveWindow.Selection.ShapeRange(1)
        If VTIsVisualTeXPowerPointShape(selectedShape) Then
            VTEnsurePowerPointFormulaScaleState _
                selectedShape, selectedSize, referenceWidthPt, referenceHeightPt
        ElseIf selectedShape.HasTextFrame = msoTrue Then
            If selectedShape.TextFrame.HasText = msoTrue Then
                selectedSize = selectedShape.TextFrame.TextRange.Font.Size
            End If
        End If
    End If

UseDefault:
    If Not VTValidPowerPointFormulaFontSize(selectedSize) Then
        selectedSize = VT_POWERPOINT_DEFAULT_FONT_SIZE_PT
    End If
    VTPreferredPowerPointFormulaFontSize = selectedSize
End Function

Private Sub VTApplyPowerPointFormulaFontSize( _
    ByVal target As Shape, _
    ByVal requestedFontSizePt As Double)

    Dim storedFontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double
    Dim targetWidth As Double
    Dim targetHeight As Double
    Dim centerX As Double
    Dim centerY As Double
    Dim originalLeft As Single
    Dim originalTop As Single
    Dim originalWidth As Single
    Dim originalHeight As Single
    Dim resizeErrorNumber As Long
    Dim resizeErrorDescription As String

    If target Is Nothing Or _
       Not VTValidPowerPointFormulaFontSize(requestedFontSizePt) Then
        Err.Raise vbObjectError + 7525, "VisualTeX", _
            "Enter a VisualTeX PowerPoint formula font size from 1 to 512 pt."
    End If
    VTEnsurePowerPointFormulaScaleState _
        target, storedFontSizePt, referenceWidthPt, referenceHeightPt
    targetWidth = referenceWidthPt * requestedFontSizePt / _
        VT_POWERPOINT_REFERENCE_FONT_SIZE_PT
    targetHeight = referenceHeightPt * requestedFontSizePt / _
        VT_POWERPOINT_REFERENCE_FONT_SIZE_PT
    If targetWidth <= 0# Or targetHeight <= 0# Or _
       targetWidth > 10000# Or targetHeight > 10000# Then
        Err.Raise vbObjectError + 7526, "VisualTeX", _
            "The requested PowerPoint formula font size produces unsupported dimensions."
    End If

    originalLeft = target.Left
    originalTop = target.Top
    originalWidth = target.Width
    originalHeight = target.Height
    centerX = CDbl(originalLeft) + CDbl(originalWidth) / 2#
    centerY = CDbl(originalTop) + CDbl(originalHeight) / 2#
    On Error GoTo ResizeFailed
    target.LockAspectRatio = msoFalse
    target.Width = CSng(targetWidth)
    target.Height = CSng(targetHeight)
    target.Left = CSng(centerX - targetWidth / 2#)
    target.Top = CSng(centerY - targetHeight / 2#)
    target.LockAspectRatio = msoTrue
    VTSetPowerPointFormulaScaleState _
        target, requestedFontSizePt, referenceWidthPt, referenceHeightPt
    If Abs(target.Width - targetWidth) > 0.1 Or _
       Abs(target.Height - targetHeight) > 0.1 Or _
       Abs((target.Left + target.Width / 2!) - centerX) > 0.1 Or _
       Abs((target.Top + target.Height / 2!) - centerY) > 0.1 Then
        Err.Raise vbObjectError + 7527, "VisualTeX", _
            "PowerPoint did not persist the requested formula point size."
    End If
    VTInvalidatePowerPointFormulaFontSizeControl
    Exit Sub

ResizeFailed:
    resizeErrorNumber = Err.Number
    resizeErrorDescription = Err.Description
    On Error Resume Next
    target.LockAspectRatio = msoFalse
    target.Left = originalLeft
    target.Top = originalTop
    target.Width = originalWidth
    target.Height = originalHeight
    target.LockAspectRatio = msoTrue
    VTSetPowerPointFormulaScaleState _
        target, storedFontSizePt, referenceWidthPt, referenceHeightPt
    On Error GoTo 0
    Err.Raise resizeErrorNumber, "VisualTeX PowerPoint formula font size", _
        resizeErrorDescription
End Sub

Public Sub VisualTeX_SynchronizeSelectedFormulaSize( _
    ByVal selected As PowerPoint.Selection)

    Dim selectedShape As Shape
    Dim fontSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double

    ' Ribbon invalidation can synchronously cause another PowerPoint selection
    ' notification while the original event is still on the VBA stack. Without
    ' a guard, loading the add-in can recurse indefinitely before the first
    ' presentation window appears.
    If VT_POWERPOINT_SELECTION_SYNC_ACTIVE Then Exit Sub
    VT_POWERPOINT_SELECTION_SYNC_ACTIVE = True
    On Error GoTo SynchronizeFinished
    If selected Is Nothing Then GoTo SynchronizeFinished
    If selected.Type <> ppSelectionShapes Or _
       selected.ShapeRange.Count <> 1 Then GoTo SynchronizeFinished
    Set selectedShape = selected.ShapeRange(1)
    If Not VTIsVisualTeXPowerPointShape(selectedShape) Then GoTo SynchronizeFinished
    VTEnsurePowerPointFormulaScaleState _
        selectedShape, fontSizePt, referenceWidthPt, referenceHeightPt

SynchronizeFinished:
    On Error Resume Next
    VTInvalidatePowerPointFormulaFontSizeControl
    VT_POWERPOINT_SELECTION_SYNC_ACTIVE = False
    On Error GoTo 0
End Sub

Public Sub VisualTeX_SetSelectedFormulaFontSize( _
    ByVal enteredText As String)

    Dim selectedShape As Shape
    Dim requestedFontSizePt As Double

    If Not VTTryParsePowerPointFormulaFontSize( _
       enteredText, requestedFontSizePt) Then
        Err.Raise vbObjectError + 7528, "VisualTeX", _
            "Enter a formula font size such as 18, 24, 28 or 32."
    End If
    Set selectedShape = VTSelectedSingleShape()
    If Not VTIsVisualTeXPowerPointShape(selectedShape) Then
        Err.Raise vbObjectError + 7529, "VisualTeX", _
            "Select one VisualTeX PowerPoint SVG formula first."
    End If
    VTApplyPowerPointFormulaFontSize selectedShape, requestedFontSizePt
End Sub

Private Function VTTryParsePowerPointFormulaFontSize( _
    ByVal enteredText As String, _
    ByRef fontSizePt As Double) As Boolean

    Dim normalized As String

    fontSizePt = 0#
    normalized = Trim$(enteredText)
    If Len(normalized) = 0 Or Len(normalized) > 32 Then Exit Function
    On Error GoTo InvalidSize
    fontSizePt = VTParseInvariantDouble(normalized)
    VTTryParsePowerPointFormulaFontSize = _
        VTValidPowerPointFormulaFontSize(fontSizePt)
    Exit Function

InvalidSize:
    fontSizePt = 0#
    Err.Clear
End Function

Public Sub VTInvalidatePowerPointFormulaFontSizeControl()
    On Error Resume Next
    If Not VT_POWERPOINT_RIBBON Is Nothing Then
        VT_POWERPOINT_RIBBON.InvalidateControl VT_POWERPOINT_FONT_CONTROL_ID
    End If
    On Error GoTo 0
End Sub

Private Function VTPowerPointSelectedFormulaFontState( _
    ByRef formulaCount As Long, _
    ByRef mixedSizes As Boolean, _
    ByRef fontSizePt As Double) As Boolean

    Dim candidate As Shape
    Dim candidateSizePt As Double
    Dim referenceWidthPt As Double
    Dim referenceHeightPt As Double

    formulaCount = 0
    mixedSizes = False
    fontSizePt = 0#
    On Error GoTo StateFailed
    If ActiveWindow Is Nothing Then Exit Function
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Function

    For Each candidate In ActiveWindow.Selection.ShapeRange
        If VTIsVisualTeXPowerPointShape(candidate) Then
            VTEnsurePowerPointFormulaScaleState _
                candidate, candidateSizePt, referenceWidthPt, _
                referenceHeightPt
            formulaCount = formulaCount + 1
            If formulaCount = 1 Then
                fontSizePt = candidateSizePt
            ElseIf Abs(candidateSizePt - fontSizePt) > 0.05 Then
                mixedSizes = True
            End If
        End If
    Next candidate

    VTPowerPointSelectedFormulaFontState = (formulaCount > 0)
    Exit Function

StateFailed:
    formulaCount = 0
    mixedSizes = False
    fontSizePt = 0#
    Err.Clear
End Function

Private Sub VTApplyPowerPointFontSizePresetToSelection( _
    ByVal requestedFontSizePt As Double)

    Dim candidate As Shape
    Dim appliedCount As Long

    If Not VTValidPowerPointFormulaFontSize(requestedFontSizePt) Then
        Err.Raise vbObjectError + 7530, "VisualTeX", _
            "The selected SVG formula font-size preset is invalid."
    End If
    If ActiveWindow Is Nothing Or _
       ActiveWindow.Selection.Type <> ppSelectionShapes Then
        Err.Raise vbObjectError + 7529, "VisualTeX", _
            "Select one or more VisualTeX PowerPoint SVG formulas first."
    End If

    For Each candidate In ActiveWindow.Selection.ShapeRange
        If VTIsVisualTeXPowerPointShape(candidate) Then
            VTApplyPowerPointFormulaFontSize _
                candidate, requestedFontSizePt
            appliedCount = appliedCount + 1
        End If
    Next candidate
    If appliedCount = 0 Then
        Err.Raise vbObjectError + 7529, "VisualTeX", _
            "Select one or more VisualTeX PowerPoint SVG formulas first."
    End If
    VTInvalidatePowerPointFormulaFontSizeControl
End Sub

Public Sub VTPowerPointRibbonGetFormulaFontSizeItemCount( _
    ByVal control As IRibbonControl, _
    ByRef returnedValue)

    returnedValue = VTFormulaFontPresetCount() + 1
End Sub

Public Sub VTPowerPointRibbonGetFormulaFontSizeItemLabel( _
    ByVal control As IRibbonControl, _
    ByVal itemIndex As Integer, _
    ByRef returnedValue)

    Dim formulaCount As Long
    Dim mixedSizes As Boolean
    Dim fontSizePt As Double

    On Error GoTo LabelFailed
    If itemIndex = 0 Then
        If Not VTPowerPointSelectedFormulaFontState( _
           formulaCount, mixedSizes, fontSizePt) Then
            returnedValue = _
                VTUnicodeText(35831, 36873, 25321) & " SVG " & _
                VTUnicodeText(20844, 24335)
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

Public Sub VTPowerPointRibbonGetFormulaFontSizeSelectedIndex( _
    ByVal control As IRibbonControl, _
    ByRef returnedValue)

    Dim formulaCount As Long
    Dim mixedSizes As Boolean
    Dim fontSizePt As Double
    Dim presetIndex As Long

    returnedValue = 0
    On Error GoTo SelectedFinished
    If Not VTPowerPointSelectedFormulaFontState( _
       formulaCount, mixedSizes, fontSizePt) Then Exit Sub
    If mixedSizes Then Exit Sub
    presetIndex = VTFormulaFontPresetIndex(fontSizePt)
    If presetIndex >= 0 Then returnedValue = presetIndex + 1

SelectedFinished:
End Sub

Public Sub VTPowerPointRibbonApplyFormulaFontSizePreset( _
    ByVal control As IRibbonControl, _
    ByVal selectedId As String, _
    ByVal selectedIndex As Integer)

    On Error GoTo Failed
    If selectedIndex <= 0 Then Exit Sub
    VTApplyPowerPointFontSizePresetToSelection _
        VTFormulaFontPresetSize(selectedIndex - 1)
    Exit Sub

Failed:
    VTShowError "PowerPoint SVG formula font size", _
        Err.Number, Err.Description
End Sub

Private Function VTPowerPointGeometryJson( _
    ByVal currentSlide As Slide, _
    ByVal target As Shape, _
    Optional ByVal fontSizePt As Double = 0#, _
    Optional ByVal referenceWidthPt As Double = 0#, _
    Optional ByVal referenceHeightPt As Double = 0#) As String

    Dim shapeIndex As Long

    shapeIndex = VTPowerPointShapeCollectionIndex(currentSlide, target)
    If shapeIndex <= 0 Then
        Err.Raise vbObjectError + 7529, "VisualTeX", _
            "The selected PowerPoint shape is no longer in its slide."
    End If

    VTPowerPointGeometryJson = "{" & _
        """presentationIdentity"":" & VTJsonString(VTPresentationIdentity()) & "," & _
        """slideIndex"":" & CStr(currentSlide.SlideIndex) & "," & _
        """slideId"":" & CStr(currentSlide.SlideID) & "," & _
        """shapeIndex"":" & CStr(shapeIndex) & "," & _
        """shapeId"":" & CStr(target.Id) & "," & _
        """shapeName"":" & VTJsonString(target.Name) & "," & _
        """left"":" & VTJsonNumber(target.Left) & "," & _
        """top"":" & VTJsonNumber(target.Top) & "," & _
        """width"":" & VTJsonNumber(target.Width) & "," & _
        """height"":" & VTJsonNumber(target.Height) & "," & _
        """rotation"":" & VTJsonNumber(target.Rotation) & "," & _
        """zOrder"":" & CStr(target.ZOrderPosition) & "," & _
        """fontSizePt"":" & _
            IIf(fontSizePt > 0#, VTJsonNumber(fontSizePt), "null") & "," & _
        """referenceWidthPt"":" & _
            IIf(referenceWidthPt > 0#, VTJsonNumber(referenceWidthPt), "null") & "," & _
        """referenceHeightPt"":" & _
            IIf(referenceHeightPt > 0#, VTJsonNumber(referenceHeightPt), "null") & _
        "}"
End Function

Private Function VTPresentationIdentityFor( _
    ByVal targetPresentation As Presentation) As String

    On Error Resume Next
    VTPresentationIdentityFor = targetPresentation.FullName
    If Err.Number <> 0 Or Len(VTPresentationIdentityFor) = 0 Then
        Err.Clear
        VTPresentationIdentityFor = targetPresentation.Name
    End If
    On Error GoTo 0
    VTPresentationIdentityFor = VTBoundedIdentity(VTPresentationIdentityFor)
End Function

Private Function VTPresentationIdentity() As String
    VTPresentationIdentity = VTPresentationIdentityFor(ActivePresentation)
End Function

Private Function VTFindPowerPointPresentation( _
    ByVal expectedIdentity As String) As Presentation

    Dim candidatePresentation As Presentation
    Dim matchingPresentation As Presentation
    Dim matchCount As Long

    expectedIdentity = VTBoundedIdentity(expectedIdentity)
    If Len(expectedIdentity) = 0 Or Presentations.Count = 0 Then
        Err.Raise vbObjectError + 7525, "VisualTeX", _
            "The target PowerPoint presentation is no longer open."
    End If

    For Each candidatePresentation In Presentations
        If VTPresentationIdentityFor(candidatePresentation) = expectedIdentity Then
            matchCount = matchCount + 1
            Set matchingPresentation = candidatePresentation
        End If
    Next candidatePresentation
    If matchCount <> 1 Or matchingPresentation Is Nothing Then
        Err.Raise vbObjectError + 7525, "VisualTeX", _
            "The target PowerPoint presentation is no longer uniquely available."
    End If

    Set VTFindPowerPointPresentation = matchingPresentation
End Function

Private Sub VTRequireWritablePowerPointPresentationObject( _
    ByVal targetPresentation As Presentation)

    If targetPresentation Is Nothing Then
        Err.Raise vbObjectError + 7510, "VisualTeX", _
            "Open the target PowerPoint presentation first."
    End If
    If targetPresentation.ReadOnly = msoTrue Then
        Err.Raise vbObjectError + 7511, "VisualTeX", _
            "The target PowerPoint presentation is read-only."
    End If
    If targetPresentation.Windows.Count = 0 Then
        Err.Raise vbObjectError + 7512, "VisualTeX", _
            "The target PowerPoint presentation has no editable window."
    End If
End Sub

Private Sub VTRequireWritablePowerPointPresentation()
    If Presentations.Count = 0 Then
        Err.Raise vbObjectError + 7510, "VisualTeX", "Open a PowerPoint presentation first."
    End If
    VTRequireWritablePowerPointPresentationObject ActivePresentation
    If ActiveWindow Is Nothing Then
        Err.Raise vbObjectError + 7512, "VisualTeX", "Switch PowerPoint to a normal editing view."
    End If
End Sub

Private Sub VTSetShapeTag(ByVal target As Shape, ByVal key As String, ByVal value As String)
    On Error Resume Next
    target.Tags.Delete key
    Err.Clear
    On Error GoTo Failed
    target.Tags.Add key, value
    Exit Sub
Failed:
    Err.Raise vbObjectError + 7513, "VisualTeX", "PowerPoint could not persist VisualTeX tag " & key & "."
End Sub

Private Sub VTRestoreZOrder(ByVal target As Shape, ByVal expectedPosition As Long)
    Dim guard As Long
    Dim currentPosition As Long
    Dim shapeCount As Long
    Dim directSteps As Long
    Dim edgeRouteSteps As Long

    If target Is Nothing Or expectedPosition <= 0 Then Exit Sub
    currentPosition = target.ZOrderPosition
    If currentPosition = expectedPosition Then Exit Sub

    On Error Resume Next
    shapeCount = target.Parent.Shapes.Count
    Err.Clear
    On Error GoTo RestoreFailed
    If shapeCount <= 0 Then shapeCount = currentPosition
    If expectedPosition > shapeCount Then expectedPosition = shapeCount

    If currentPosition > expectedPosition Then
        directSteps = currentPosition - expectedPosition
        edgeRouteSteps = expectedPosition
        If edgeRouteSteps < directSteps Then
            target.ZOrder msoSendToBack
            guard = guard + 1
            Do While target.ZOrderPosition < expectedPosition And guard < 8192
                target.ZOrder msoBringForward
                guard = guard + 1
            Loop
        Else
            Do While target.ZOrderPosition > expectedPosition And guard < 8192
                target.ZOrder msoSendBackward
                guard = guard + 1
            Loop
        End If
    Else
        directSteps = expectedPosition - currentPosition
        edgeRouteSteps = shapeCount - expectedPosition + 1
        If edgeRouteSteps < directSteps Then
            target.ZOrder msoBringToFront
            guard = guard + 1
            Do While target.ZOrderPosition > expectedPosition And guard < 8192
                target.ZOrder msoSendBackward
                guard = guard + 1
            Loop
        Else
            Do While target.ZOrderPosition < expectedPosition And guard < 8192
                target.ZOrder msoBringForward
                guard = guard + 1
            Loop
        End If
    End If

    If target.ZOrderPosition <> expectedPosition Then GoTo RestoreFailed
    Exit Sub

RestoreFailed:
    Err.Raise vbObjectError + 7514, "VisualTeX", _
        "PowerPoint could not restore the formula z-order."
End Sub

Private Function VTDispatchOptionalPpt(ByVal dispatch As Object, ByVal key As String) As String
    If VTCollectionHasKey(dispatch, key) Then VTDispatchOptionalPpt = CStr(dispatch(key))
End Function

Private Function VTDispatchDoublePpt(ByVal dispatch As Object, ByVal key As String) As Double
    VTRequireDispatchValue dispatch, key
    VTDispatchDoublePpt = VTParseInvariantDouble(CStr(dispatch(key)))
    If Abs(VTDispatchDoublePpt) > 10000000# Then
        Err.Raise vbObjectError + 7521, "VisualTeX", "VisualTeX dispatch contains invalid " & key & "."
    End If
End Function

Private Function VTDispatchPositiveDoublePpt(ByVal dispatch As Object, ByVal key As String) As Double
    VTDispatchPositiveDoublePpt = VTDispatchDoublePpt(dispatch, key)
    If VTDispatchPositiveDoublePpt <= 0# Then
        Err.Raise vbObjectError + 7522, "VisualTeX", "VisualTeX dispatch contains invalid " & key & "."
    End If
End Function

Private Sub VTWritePowerPointHealth()
    Dim statusPath As String
    Dim payload As String
    statusPath = VTApplicationSupportRoot() & VT_POWERPOINT_STATUS_FILE
    payload = "{" & _
        """loaded"":true," & _
        """pluginVersion"":" & VTJsonString(VT_PLUGIN_VERSION) & "," & _
        """sourceRevision"":" & _
            VTJsonString(VT_POWERPOINT_SOURCE_REVISION) & "," & _
        """host"":""powerpoint""," & _
        """timestamp"":" & VTJsonString(Format$(Now, "yyyy-mm-dd\Thh:nn:ss")) & _
        "}"
    VTWriteTextAtomic statusPath, payload
End Sub
