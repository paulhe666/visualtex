Attribute VB_Name = "VTWordAdapter"
Option Explicit

Private Const VT_WORD_HOST As String = "word"
Private Const VT_WORD_STATUS_FILE As String = "/OfficePluginStatus/word.json"
Private Const VT_WORD_BOOKMARK_PREFIX As String = "VT_Pending_"
Private Const VT_WORD_LATEX_VARIABLE_PREFIX As String = "VT_Latex_"
Private Const VT_WORD_LATEX_CHUNK_SIZE As Long = 20000
Private Const VT_WORD_LATEX_MAX_CHUNKS As Long = 128
Private VT_WORD_EVENT_SINK As VTWordEvents

Public Sub AutoExec()
    On Error Resume Next
    VTInitializeWordEvents
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

Public Sub VTInitializeWordEvents()
    Set VT_WORD_EVENT_SINK = New VTWordEvents
    Set VT_WORD_EVENT_SINK.App = Word.Application
End Sub

Public Sub VisualTeX_CreateInline()
    VTWordCreate "inline", False
End Sub

Public Sub VisualTeX_CreateDisplay()
    VTWordCreate "block", False
End Sub

Public Sub VisualTeX_CreateNumberedDisplay()
    VTWordCreate "block", True
End Sub

Public Sub VisualTeX_EditSelected()
    On Error GoTo Failed

    VTRequireWritableWordDocument
    If Selection.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7400, "VisualTeX", "Select exactly one VisualTeX inline formula."
    End If

    VTWordEditInlineShape Selection.InlineShapes(1)
    Exit Sub

Failed:
    VTShowError "Word edit", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_EditInlineShape(ByVal selectedShape As InlineShape)
    On Error GoTo Failed
    VTRequireWritableWordDocument
    VTWordEditInlineShape selectedShape
    Exit Sub
Failed:
    VTShowError "Word edit", Err.Number, Err.Description
End Sub

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

Private Sub VTWordEditInlineShape(ByVal selectedShape As InlineShape)
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
        "")
    VTWriteRequest sessionId, requestJson
    VTLaunchSession VT_WORD_HOST, sessionId
End Sub

Public Sub VisualTeX_ConvertSelectedToNativeEquation()
    On Error GoTo Failed

    VTRequireWritableWordDocument
    If Selection.InlineShapes.Count <> 1 Then
        Err.Raise vbObjectError + 7428, "VisualTeX", "Select exactly one VisualTeX formula image to convert."
    End If
    VTWordConvertInlineShapeToNativeEquation Selection.InlineShapes(1)
    Exit Sub

Failed:
    VTShowError "Word native equation conversion", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_UpdateEquationNumbers()
    On Error GoTo Failed
    Dim field As Field
    Dim updated As Long
    Dim equationLabelName As String

    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7401, "VisualTeX", "Open a Word document first."
    End If
    equationLabelName = VTNativeEquationLabelName()
    For Each field In ActiveDocument.Fields
        If VTIsNativeEquationSequenceField(field, equationLabelName) Then
            field.Update
            updated = updated + 1
        End If
    Next field
    VTShowInformation "Updated " & CStr(updated) & " native Word equation numbers."
    Exit Sub

Failed:
    VTShowError "equation numbering", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_OpenApplication()
    On Error GoTo Failed
    VTOpenApplication VT_WORD_HOST
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

    sessionId = VTReadActiveSessionId(VT_WORD_HOST)
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

Private Sub VTWordCreate(ByVal displayMode As String, ByVal numbered As Boolean)
    On Error GoTo Failed

    Dim sessionId As String
    Dim formulaId As String
    Dim pendingMarker As String
    Dim placeholder As InlineShape
    Dim insertionRange As Range
    Dim requestJson As String
    Dim errorNumber As Long
    Dim errorDescription As String

    VTRequireWritableWordDocument
    If Not VTPathFileExists(VTPlaceholderImagePath()) Then
        Err.Raise vbObjectError + 7404, "VisualTeX", "The VisualTeX placeholder resource is missing. Repair the offline add-in."
    End If

    sessionId = VTNewUuidV4()
    formulaId = VTNewUuidV4()
    pendingMarker = VTPendingMarker(sessionId, formulaId)
    Set insertionRange = Selection.Range.Duplicate
    insertionRange.Collapse wdCollapseStart
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
        "")
    VTWriteRequest sessionId, requestJson
    VTLaunchSession VT_WORD_HOST, sessionId
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
    Dim imagePath As String
    Dim metadata As String
    Dim latexBase64 As String
    Dim previousLatexBase64 As String
    Dim hadPreviousLatexPayload As Boolean
    Dim latexPayloadStored As Boolean
    Dim pendingMarker As String
    Dim sourceMarker As String
    Dim sourceDocumentId As String
    Dim target As InlineShape
    Dim committed As InlineShape
    Dim candidate As InlineShape
    Dim insertionRange As Range
    Dim widthPoints As Double
    Dim heightPoints As Double
    Dim baselinePoints As Double
    Dim formulaReference As String
    Dim insertedNumber As Range

    VTRequireWritableWordDocument
    VTRequireDispatchValue dispatch, "mode"
    VTRequireDispatchValue dispatch, "formulaId"
    VTRequireDispatchValue dispatch, "displayMode"
    VTRequireDispatchValue dispatch, "numbered"
    VTRequireDispatchValue dispatch, "imagePath"
    VTRequireDispatchValue dispatch, "metadata"
    VTRequireDispatchValue dispatch, "latexBase64"

    mode = CStr(dispatch("mode"))
    formulaId = CStr(dispatch("formulaId"))
    displayMode = CStr(dispatch("displayMode"))
    numbered = (CStr(dispatch("numbered")) = "1")
    imagePath = CStr(dispatch("imagePath"))
    metadata = CStr(dispatch("metadata"))
    latexBase64 = CStr(dispatch("latexBase64"))
    pendingMarker = VTDispatchOptional(dispatch, "pendingMarker")
    sourceMarker = VTDispatchOptional(dispatch, "sourceMarker")
    sourceDocumentId = VTDispatchOptional(dispatch, "sourceDocumentId")
    If Len(sourceDocumentId) = 0 Or sourceDocumentId <> VTWordDocumentIdentity() Then
        Err.Raise vbObjectError + 7415, "VisualTeX", "The active Word document changed while VisualTeX was open."
    End If
    widthPoints = VTDispatchPositiveDouble(dispatch, "widthPoints")
    heightPoints = VTDispatchPositiveDouble(dispatch, "heightPoints")
    baselinePoints = VTDispatchOptionalDouble(dispatch, "baseline", 0#)

    If Not VTIsCanonicalUuid(formulaId) Or Not VTIsEncodedMetadata(metadata) Or _
       Not VTIsBase64UrlPayload(latexBase64) Then
        Err.Raise vbObjectError + 7405, "VisualTeX", "VisualTeX Word result metadata or native-equation payload is invalid."
    End If
    formulaReference = VTFormulaReference(formulaId, displayMode, numbered)
    VTValidateAbsoluteVisualTeXPath imagePath
    If Not VTPathFileExists(imagePath) Then
        Err.Raise vbObjectError + 7406, "VisualTeX", "VisualTeX Word result image is missing."
    End If

    On Error Resume Next
    If mode = "create" Then
        Set target = VTFindUniqueInlineShape(pendingMarker)
    ElseIf mode = "edit" Then
        Set target = VTFindUniqueInlineShape(sourceMarker)
    Else
        On Error GoTo 0
        Err.Raise vbObjectError + 7407, "VisualTeX", "VisualTeX Word result mode is invalid."
    End If
    Err.Clear
    On Error GoTo RollbackCandidate
    If target Is Nothing Then
        Set committed = VTFindCommittedInlineShape(metadata, formulaReference)
        If Not committed Is Nothing Then
            VTSetWordLatexPayload ActiveDocument, formulaId, latexBase64
            VTDeletePendingBookmark sessionId
            Exit Sub
        End If
        Err.Raise vbObjectError + 7426, "VisualTeX", "The original Word formula is missing and no committed VisualTeX result was found."
    End If

    Set insertionRange = target.Range.Duplicate
    insertionRange.Collapse wdCollapseStart
    On Error GoTo RollbackCandidate
    Set candidate = ActiveDocument.InlineShapes.AddPicture( _
        FileName:=imagePath, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Range:=insertionRange)
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
    Else
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
    ElseIf candidate.Range.ParagraphFormat.Alignment <> wdAlignParagraphCenter Then
        Err.Raise vbObjectError + 7424, "VisualTeX", "Word did not persist the VisualTeX display alignment."
    End If

    If mode = "create" And displayMode = "block" And numbered Then
        Set insertedNumber = VTInsertEquationNumber(candidate)
    End If

    hadPreviousLatexPayload = VTTryReadWordLatexPayload( _
        ActiveDocument, formulaId, previousLatexBase64)
    VTSetWordLatexPayload ActiveDocument, formulaId, latexBase64
    latexPayloadStored = True

    target.Delete
    VTDeletePendingBookmark sessionId
    Exit Sub

RollbackCandidate:
    Dim transactionErrorNumber As Long
    Dim transactionErrorDescription As String
    transactionErrorNumber = Err.Number
    transactionErrorDescription = Err.Description
    On Error Resume Next
    If latexPayloadStored Then
        If hadPreviousLatexPayload Then
            VTSetWordLatexPayload ActiveDocument, formulaId, previousLatexBase64
        Else
            VTDeleteWordLatexPayload ActiveDocument, formulaId
        End If
    End If
    If Not insertedNumber Is Nothing Then insertedNumber.Delete
    If Not candidate Is Nothing Then candidate.Delete
    On Error GoTo 0
    Err.Raise transactionErrorNumber, "VisualTeX Word transaction", transactionErrorDescription
End Sub

Private Sub VTCancelWordDispatch(ByVal sessionId As String, ByVal dispatch As Object)
    Dim pendingMarker As String
    Dim sourceDocumentId As String
    Dim target As InlineShape

    pendingMarker = VTDispatchOptional(dispatch, "pendingMarker")
    sourceDocumentId = VTDispatchOptional(dispatch, "sourceDocumentId")
    If Len(sourceDocumentId) > 0 And sourceDocumentId <> VTWordDocumentIdentity() Then Exit Sub
    If Len(pendingMarker) > 0 Then
        On Error Resume Next
        Set target = VTFindUniqueInlineShape(pendingMarker)
        If Not target Is Nothing Then target.Delete
        On Error GoTo 0
    End If
    VTDeletePendingBookmark sessionId
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

Private Function VTInsertEquationNumber(ByVal formulaShape As InlineShape) As Range
    Dim paragraphRange As Range
    Dim prefixRange As Range
    Dim numberRange As Range
    Dim sequenceField As Field
    Dim layoutStart As Long
    Dim numberStart As Long
    Dim textWidth As Single
    Dim equationNumberRange As Range
    Dim numberFontSize As Single
    Dim numberRaisePoints As Single
    Dim equationLabelName As String

    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range
    equationLabelName = VTNativeEquationLabelName()
    textWidth = ActiveDocument.PageSetup.TextColumns.Width
    If textWidth <= 0! Then
        Err.Raise vbObjectError + 7425, "VisualTeX", "Word returned an invalid text width for equation numbering."
    End If

    ' A centered paragraph plus one right tab treats the formula and number as
    ' one centered run, which can push the formula to the far right and wrap
    ' the number onto the next line. Use the standard three-position layout:
    ' a center tab before the formula and a right tab before the number.
    ' A built-in Caption paragraph plus a SEQ field using Word's built-in
    ' Equation label is what makes this item appear in References -> Cross-reference.
    paragraphRange.Style = wdStyleCaption
    paragraphRange.ParagraphFormat.Alignment = wdAlignParagraphLeft
    paragraphRange.ParagraphFormat.TabStops.ClearAll
    paragraphRange.ParagraphFormat.TabStops.Add _
        Position:=textWidth / 2!, _
        Alignment:=wdAlignTabCenter, _
        Leader:=wdTabLeaderSpaces
    paragraphRange.ParagraphFormat.TabStops.Add _
        Position:=textWidth, _
        Alignment:=wdAlignTabRight, _
        Leader:=wdTabLeaderSpaces

    Set prefixRange = formulaShape.Range.Duplicate
    prefixRange.Collapse wdCollapseStart
    layoutStart = prefixRange.Start
    prefixRange.InsertBefore vbTab

    numberStart = formulaShape.Range.End
    Set numberRange = formulaShape.Range.Duplicate
    numberRange.Collapse wdCollapseEnd
    numberRange.InsertAfter vbTab & "("
    numberRange.Collapse wdCollapseEnd
    Set sequenceField = ActiveDocument.Fields.Add( _
        Range:=numberRange, _
        Type:=wdFieldSequence, _
        Text:=VTEquationSequenceFieldText(equationLabelName), _
        PreserveFormatting:=False)
    sequenceField.Update
    Set numberRange = sequenceField.Result.Duplicate
    numberRange.Collapse wdCollapseEnd
    numberRange.InsertAfter ")"
    If numberRange.End <= numberStart Then
        Err.Raise vbObjectError + 7425, "VisualTeX", "Word did not create the native Equation caption number."
    End If

    ' Inline pictures sit on the paragraph baseline, so an ordinary text run
    ' beside a tall display formula appears close to the bottom of the image.
    ' Raise the complete parenthesized number by half of the difference between
    ' the formula height and the number font size to align their visual centers.
    Set equationNumberRange = formulaShape.Range.Document.Range( _
        Start:=numberStart + 1, _
        End:=numberRange.End)
    numberFontSize = equationNumberRange.Font.Size
    If numberFontSize <= 0! Or numberFontSize > 72! Then numberFontSize = 12!
    numberRaisePoints = (formulaShape.Height - numberFontSize) / 2!
    If numberRaisePoints < 0! Then numberRaisePoints = 0!
    If numberRaisePoints > 48! Then numberRaisePoints = 48!
    equationNumberRange.Font.Position = CLng(numberRaisePoints)

    Set VTInsertEquationNumber = formulaShape.Range.Document.Range( _
        Start:=layoutStart, _
        End:=numberRange.End)
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

Private Sub VTWordConvertInlineShapeToNativeEquation(ByVal target As InlineShape)
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim latexBase64 As String
    Dim latex As String
    Dim linearFormula As String
    Dim insertionRange As Range
    Dim equationRange As Range
    Dim nativeEquation As OMath
    Dim rollbackRange As Range
    Dim insertionStart As Long
    Dim candidateInserted As Boolean
    Dim conversionErrorNumber As Long
    Dim conversionErrorDescription As String

    If target Is Nothing Or Not VTIsVisualTeXInlineShape(target) Then
        Err.Raise vbObjectError + 7430, "VisualTeX", "The selected object is not a VisualTeX formula image."
    End If
    If Not VTTryParseFormulaReference(target.Title, formulaId, displayMode, numbered) Then
        Err.Raise vbObjectError + 7431, "VisualTeX", "The selected VisualTeX formula reference is invalid."
    End If
    If Not VTTryReadWordLatexPayload(ActiveDocument, formulaId, latexBase64) Then
        Err.Raise vbObjectError + 7432, "VisualTeX", _
            "This formula predates native-equation conversion metadata. Edit and save it once in VisualTeX, then convert it again."
    End If

    latex = VTBase64UrlDecodeUtf8(latexBase64)
    linearFormula = VTLaTeXToWordLinear(latex)
    If Len(Trim$(linearFormula)) = 0 Then
        Err.Raise vbObjectError + 7433, "VisualTeX", "VisualTeX could not produce a Word linear equation from the stored LaTeX."
    End If

    insertionStart = target.Range.Start
    Set insertionRange = target.Range.Duplicate
    insertionRange.Collapse wdCollapseStart
    insertionRange.Text = linearFormula
    insertionRange.SetRange Start:=insertionStart, End:=insertionStart + Len(linearFormula)
    candidateInserted = True

    On Error GoTo RollbackConversion
    Set equationRange = ActiveDocument.OMaths.Add(insertionRange)
    If equationRange.OMaths.Count <> 1 Then
        Err.Raise vbObjectError + 7434, "VisualTeX", "Word did not create exactly one native equation object."
    End If
    Set nativeEquation = equationRange.OMaths(1)
    nativeEquation.BuildUp
    equationRange.Font.Position = 0

    If displayMode = "inline" Or numbered Then
        ' A numbered display formula must remain in the same tabbed paragraph
        ' as its native Equation caption field, so it uses inline OMath layout.
        nativeEquation.Type = wdOMathInline
    Else
        nativeEquation.Type = wdOMathDisplay
        equationRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If

    target.Delete
    On Error Resume Next
    VTDeleteWordLatexPayload ActiveDocument, formulaId
    equationRange.Select
    On Error GoTo 0
    Exit Sub

RollbackConversion:
    conversionErrorNumber = Err.Number
    conversionErrorDescription = Err.Description
    On Error Resume Next
    If candidateInserted And Not target Is Nothing Then
        Set rollbackRange = target.Range.Document.Range( _
            Start:=insertionStart, _
            End:=target.Range.Start)
        If rollbackRange.End > rollbackRange.Start Then rollbackRange.Delete
    End If
    On Error GoTo 0
    Err.Raise conversionErrorNumber, "VisualTeX Word native equation conversion", conversionErrorDescription
End Sub

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

Private Sub VTAddPendingBookmark(ByVal targetRange As Range, ByVal sessionId As String)
    Dim name As String
    name = VTWordBookmarkName(sessionId)
    On Error Resume Next
    If ActiveDocument.Bookmarks.Exists(name) Then ActiveDocument.Bookmarks(name).Delete
    On Error GoTo 0
    ActiveDocument.Bookmarks.Add Name:=name, Range:=targetRange
End Sub

Private Sub VTDeletePendingBookmark(ByVal sessionId As String)
    Dim name As String
    name = VTWordBookmarkName(sessionId)
    On Error Resume Next
    If ActiveDocument.Bookmarks.Exists(name) Then ActiveDocument.Bookmarks(name).Delete
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
