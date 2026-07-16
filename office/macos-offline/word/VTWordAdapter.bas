Attribute VB_Name = "VTWordAdapter"
Option Explicit

Private Const VT_WORD_HOST As String = "word"
Private Const VT_WORD_STATUS_FILE As String = "/OfficePluginStatus/word.json"
Private Const VT_WORD_BOOKMARK_PREFIX As String = "VT_Pending_"
Private Const VT_WORD_SEQUENCE_NAME As String = "VisualTeXEquation"
Private VT_WORD_EVENT_SINK As VTWordEvents

Public Sub AutoExec()
    On Error Resume Next
    VTInitializeWordEvents
    VTWriteWordHealth
    On Error GoTo 0
End Sub

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

Public Sub VisualTeX_UpdateEquationNumbers()
    On Error GoTo Failed
    Dim field As Field
    Dim updated As Long

    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7401, "VisualTeX", "Open a Word document first."
    End If
    For Each field In ActiveDocument.Fields
        If field.Type = wdFieldSequence Then
            If InStr(1, field.Code.Text, VT_WORD_SEQUENCE_NAME, vbTextCompare) > 0 Then
                field.Update
                updated = updated + 1
            End If
        End If
    Next field
    VTShowInformation "Updated " & CStr(updated) & " VisualTeX equation numbers."
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

    mode = CStr(dispatch("mode"))
    formulaId = CStr(dispatch("formulaId"))
    displayMode = CStr(dispatch("displayMode"))
    numbered = (CStr(dispatch("numbered")) = "1")
    imagePath = CStr(dispatch("imagePath"))
    metadata = CStr(dispatch("metadata"))
    pendingMarker = VTDispatchOptional(dispatch, "pendingMarker")
    sourceMarker = VTDispatchOptional(dispatch, "sourceMarker")
    sourceDocumentId = VTDispatchOptional(dispatch, "sourceDocumentId")
    If Len(sourceDocumentId) = 0 Or sourceDocumentId <> VTWordDocumentIdentity() Then
        Err.Raise vbObjectError + 7415, "VisualTeX", "The active Word document changed while VisualTeX was open."
    End If
    widthPoints = VTDispatchPositiveDouble(dispatch, "widthPoints")
    heightPoints = VTDispatchPositiveDouble(dispatch, "heightPoints")
    baselinePoints = VTDispatchOptionalDouble(dispatch, "baseline", 0#)

    If Not VTIsCanonicalUuid(formulaId) Or Not VTIsEncodedMetadata(metadata) Then
        Err.Raise vbObjectError + 7405, "VisualTeX", "VisualTeX Word result metadata is invalid."
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

    target.Delete
    VTDeletePendingBookmark sessionId
    Exit Sub

RollbackCandidate:
    Dim transactionErrorNumber As Long
    Dim transactionErrorDescription As String
    transactionErrorNumber = Err.Number
    transactionErrorDescription = Err.Description
    On Error Resume Next
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

    Set paragraphRange = formulaShape.Range.Paragraphs(1).Range
    textWidth = ActiveDocument.PageSetup.TextColumns.Width
    If textWidth <= 0! Then
        Err.Raise vbObjectError + 7425, "VisualTeX", "Word returned an invalid text width for equation numbering."
    End If

    ' A centered paragraph plus one right tab treats the formula and number as
    ' one centered run, which can push the formula to the far right and wrap
    ' the number onto the next line. Use the standard three-position layout:
    ' a center tab before the formula and a right tab before the number.
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
        Text:=VT_WORD_SEQUENCE_NAME & " \* ARABIC", _
        PreserveFormatting:=False)
    sequenceField.Update
    Set numberRange = sequenceField.Result.Duplicate
    numberRange.Collapse wdCollapseEnd
    numberRange.InsertAfter ")"
    If numberRange.End <= numberStart Then
        Err.Raise vbObjectError + 7425, "VisualTeX", "Word did not create the VisualTeX equation number."
    End If
    Set VTInsertEquationNumber = formulaShape.Range.Document.Range( _
        Start:=layoutStart, _
        End:=numberRange.End)
End Function

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
