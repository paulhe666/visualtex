Attribute VB_Name = "VTPowerPointAdapter"
Option Explicit

Private Const VT_POWERPOINT_HOST As String = "powerpoint"
Private Const VT_POWERPOINT_STATUS_FILE As String = "/OfficePluginStatus/powerpoint.json"
Private Const VT_SHAPE_PREFIX As String = "VisualTeX_"
Private Const VT_DEFAULT_PLACEHOLDER_WIDTH As Single = 180!
Private Const VT_DEFAULT_PLACEHOLDER_HEIGHT As Single = 42!
Private VT_POWERPOINT_EVENT_SINK As VTPowerPointEvents

Public Sub Auto_Open()
    On Error Resume Next
    VTInitializePowerPointEvents
    VTWritePowerPointHealth
    On Error GoTo 0
End Sub

Public Sub VTInitializePowerPointEvents()
    Set VT_POWERPOINT_EVENT_SINK = New VTPowerPointEvents
    Set VT_POWERPOINT_EVENT_SINK.App = PowerPoint.Application
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
    Dim failureStage As String

    failureStage = "validate presentation"
    VTRequireWritablePowerPointPresentation
    Set currentSlide = ActiveWindow.View.Slide

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
    VTSetShapeTag placeholder, "VisualTeXSessionId", sessionId
    VTSetShapeTag placeholder, "VisualTeXPending", "1"
    placeholder.AlternativeText = pendingMarker

    failureStage = "build request"
    powerPointJson = VTPowerPointGeometryJson(currentSlide, placeholder)
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

    failureStage = "write request"
    VTWriteRequest sessionId, requestJson

    failureStage = "open VisualTeX editor"
    VTLaunchSession VT_POWERPOINT_HOST, sessionId
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

Private Sub VTPowerPointEditShape(ByVal selectedShape As Shape)
    Dim formulaId As String
    Dim encodedMetadata As String
    Dim formulaReference As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim sessionId As String
    Dim requestJson As String
    Dim powerPointJson As String

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

    sessionId = VTNewUuidV4()
    powerPointJson = VTPowerPointGeometryJson(ActiveWindow.View.Slide, selectedShape)
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
        powerPointJson)
    VTWriteRequest sessionId, requestJson
    VTLaunchSession VT_POWERPOINT_HOST, sessionId
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
    Dim currentSlide As Slide
    Dim committed As Shape
    Dim original As Shape
    Dim candidate As Shape
    Dim originalTemporaryName As String
    Dim candidateTemporaryName As String
    Dim originalRenamed As Boolean
    Dim formulaReference As String

    VTRequireWritablePowerPointPresentation
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
    VTRequireDispatchValue dispatch, "rotation"
    VTRequireDispatchValue dispatch, "zOrder"

    formulaId = CStr(dispatch("formulaId"))
    metadata = CStr(dispatch("metadata"))
    shapeName = CStr(dispatch("shapeName"))
    sourceShapeName = CStr(dispatch("sourceShapeName"))
    imagePath = CStr(dispatch("imagePath"))
    fallbackImagePath = VTDispatchOptionalPpt(dispatch, "fallbackImagePath")
    expectedPresentation = CStr(dispatch("presentationIdentity"))
    slideIndex = CLng(dispatch("slideIndex"))
    slideId = CLng(dispatch("slideId"))
    targetLeft = VTDispatchDoublePpt(dispatch, "targetLeft")
    targetTop = VTDispatchDoublePpt(dispatch, "targetTop")
    targetWidth = VTDispatchPositiveDoublePpt(dispatch, "targetWidth")
    targetHeight = VTDispatchPositiveDoublePpt(dispatch, "targetHeight")
    targetRotation = VTDispatchDoublePpt(dispatch, "rotation")
    targetZOrder = CLng(dispatch("zOrder"))

    If Not VTIsCanonicalUuid(formulaId) Or shapeName <> VT_SHAPE_PREFIX & formulaId Or Not VTIsEncodedMetadata(metadata) Then
        Err.Raise vbObjectError + 7504, "VisualTeX", "VisualTeX PowerPoint result metadata is invalid."
    End If
    formulaReference = VTFormulaReference(formulaId, "block", False)
    If expectedPresentation <> VTPresentationIdentity() Then
        Err.Raise vbObjectError + 7515, "VisualTeX", "The active PowerPoint presentation changed while VisualTeX was open."
    End If
    If slideIndex <= 0 Or slideIndex > ActivePresentation.Slides.Count Or slideId <= 0 Or targetZOrder <= 0 Then
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

    Set currentSlide = ActivePresentation.Slides(slideIndex)
    If currentSlide.SlideID <> slideId Then
        Err.Raise vbObjectError + 7518, "VisualTeX", "The original PowerPoint slide no longer exists."
    End If
    On Error Resume Next
    Set committed = currentSlide.Shapes(shapeName)
    On Error GoTo TransactionFailed
    If Not committed Is Nothing Then
        If VTIsCommittedPowerPointShape( _
            committed, shapeName, formulaReference, metadata, formulaId, sessionId, _
            targetLeft, targetTop, targetWidth, targetHeight, targetRotation, targetZOrder) Then
            Exit Sub
        End If
    End If
    Set committed = Nothing

    On Error Resume Next
    Set original = currentSlide.Shapes(sourceShapeName)
    On Error GoTo TransactionFailed
    If original Is Nothing Then
        Err.Raise vbObjectError + 7519, "VisualTeX", "The original VisualTeX PowerPoint shape no longer exists."
    End If

    candidateTemporaryName = "VisualTeXPendingResult_" & Replace$(Left$(sessionId, 13), "-", "")
    originalTemporaryName = "VisualTeXOriginal_" & Replace$(Left$(sessionId, 13), "-", "")
    ' Modern PowerPoint for Mac preserves an imported SVG as vector artwork.
    ' Prefer SVG so formulas remain sharp at arbitrary zoom. Keep the PNG only
    ' as a compatibility fallback for Office builds that reject SVG AddPicture.
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
    On Error GoTo TransactionFailed
    If candidate Is Nothing Then
        If Len(fallbackImagePath) = 0 Then
            Err.Raise vbObjectError + 7521, "VisualTeX", _
                "PowerPoint could not insert the VisualTeX SVG: " & _
                CStr(vectorInsertErrorNumber) & " " & vectorInsertErrorDescription
        End If
        Set candidate = currentSlide.Shapes.AddPicture( _
            FileName:=fallbackImagePath, _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=CSng(targetLeft), _
            Top:=CSng(targetTop), _
            Width:=CSng(targetWidth), _
            Height:=CSng(targetHeight))
    End If
    candidate.Name = candidateTemporaryName
    candidate.LockAspectRatio = msoFalse
    candidate.Left = CSng(targetLeft)
    candidate.Top = CSng(targetTop)
    candidate.Width = CSng(targetWidth)
    candidate.Height = CSng(targetHeight)
    candidate.LockAspectRatio = msoTrue
    candidate.Rotation = CSng(targetRotation)
    candidate.AlternativeText = metadata
    candidate.Title = formulaReference
    VTSetShapeTag candidate, "VisualTeXFormulaId", formulaId
    VTSetShapeTag candidate, "VisualTeXSessionId", sessionId
    VTSetShapeTag candidate, "VisualTeXPending", "0"
    On Error Resume Next
    VTSetShapeTag candidate, "VisualTeXMetadata", metadata
    Err.Clear
    On Error GoTo TransactionFailed

    ' The original still occupies targetZOrder. Put the candidate immediately
    ' above it; deleting the original as the final mutation shifts the candidate
    ' into the exact original z-order without any fallible operation afterwards.
    VTRestoreZOrder candidate, targetZOrder + 1
    original.Name = originalTemporaryName
    originalRenamed = True
    candidate.Name = shapeName

    If candidate.Name <> shapeName Or _
       Abs(candidate.Left - targetLeft) > 0.1 Or Abs(candidate.Top - targetTop) > 0.1 Or _
       Abs(candidate.Width - targetWidth) > 0.1 Or Abs(candidate.Height - targetHeight) > 0.1 Or _
       Abs(candidate.Rotation - targetRotation) > 0.1 Or candidate.ZOrderPosition <> targetZOrder + 1 Or _
       candidate.AlternativeText <> metadata Or candidate.Title <> formulaReference Or _
       candidate.Tags("VisualTeXFormulaId") <> formulaId Or _
       candidate.Tags("VisualTeXSessionId") <> sessionId Or _
       candidate.Tags("VisualTeXPending") <> "0" Then
        Err.Raise vbObjectError + 7520, "VisualTeX", "PowerPoint did not persist the VisualTeX formula properties."
    End If

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
    ByVal targetZOrder As Long) As Boolean

    On Error GoTo NotCommitted
    VTIsCommittedPowerPointShape = _
        target.Name = expectedName And _
        Abs(target.Left - targetLeft) <= 0.1 And _
        Abs(target.Top - targetTop) <= 0.1 And _
        Abs(target.Width - targetWidth) <= 0.1 And _
        Abs(target.Height - targetHeight) <= 0.1 And _
        Abs(target.Rotation - targetRotation) <= 0.1 And _
        target.ZOrderPosition = targetZOrder And _
        target.AlternativeText = metadata And _
        target.Title = formulaReference And _
        target.Tags("VisualTeXFormulaId") = formulaId And _
        target.Tags("VisualTeXSessionId") = sessionId And _
        target.Tags("VisualTeXPending") = "0"
    Exit Function

NotCommitted:
    Err.Clear
    VTIsCommittedPowerPointShape = False
End Function

Private Sub VTCancelPowerPointDispatch(ByVal sessionId As String, ByVal dispatch As Object)
    Dim pendingMarker As String
    Dim currentSlide As Slide
    Dim shapeItem As Shape

    pendingMarker = VTDispatchOptionalPpt(dispatch, "pendingMarker")
    If Len(pendingMarker) = 0 Or Presentations.Count = 0 Then Exit Sub
    On Error Resume Next
    For Each currentSlide In ActivePresentation.Slides
        For Each shapeItem In currentSlide.Shapes
            If shapeItem.AlternativeText = pendingMarker And _
               shapeItem.Tags("VisualTeXSessionId") = sessionId And _
               shapeItem.Tags("VisualTeXPending") = "1" Then
                shapeItem.Delete
                Exit Sub
            End If
        Next shapeItem
    Next currentSlide
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

Private Function VTPowerPointGeometryJson(ByVal currentSlide As Slide, ByVal target As Shape) As String
    VTPowerPointGeometryJson = "{" & _
        """presentationIdentity"":" & VTJsonString(VTPresentationIdentity()) & "," & _
        """slideIndex"":" & CStr(currentSlide.SlideIndex) & "," & _
        """slideId"":" & CStr(currentSlide.SlideID) & "," & _
        """shapeName"":" & VTJsonString(target.Name) & "," & _
        """left"":" & VTJsonNumber(target.Left) & "," & _
        """top"":" & VTJsonNumber(target.Top) & "," & _
        """width"":" & VTJsonNumber(target.Width) & "," & _
        """height"":" & VTJsonNumber(target.Height) & "," & _
        """rotation"":" & VTJsonNumber(target.Rotation) & "," & _
        """zOrder"":" & CStr(target.ZOrderPosition) & _
        "}"
End Function

Private Function VTPresentationIdentity() As String
    On Error Resume Next
    VTPresentationIdentity = ActivePresentation.FullName
    If Err.Number <> 0 Or Len(VTPresentationIdentity) = 0 Then
        Err.Clear
        VTPresentationIdentity = ActivePresentation.Name
    End If
    On Error GoTo 0
    VTPresentationIdentity = VTBoundedIdentity(VTPresentationIdentity)
End Function

Private Sub VTRequireWritablePowerPointPresentation()
    If Presentations.Count = 0 Then
        Err.Raise vbObjectError + 7510, "VisualTeX", "Open a PowerPoint presentation first."
    End If
    If ActivePresentation.ReadOnly = msoTrue Then
        Err.Raise vbObjectError + 7511, "VisualTeX", "The active PowerPoint presentation is read-only."
    End If
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
    If expectedPosition <= 0 Then Exit Sub
    Do While target.ZOrderPosition > expectedPosition And guard < 4096
        target.ZOrder msoSendBackward
        guard = guard + 1
    Loop
    Do While target.ZOrderPosition < expectedPosition And guard < 8192
        target.ZOrder msoBringForward
        guard = guard + 1
    Loop
    If target.ZOrderPosition <> expectedPosition Then
        Err.Raise vbObjectError + 7514, "VisualTeX", "PowerPoint could not restore the formula z-order."
    End If
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
        """host"":""powerpoint""," & _
        """timestamp"":" & VTJsonString(Format$(Now, "yyyy-mm-dd\Thh:nn:ss")) & _
        "}"
    VTWriteTextAtomic statusPath, payload
End Sub
