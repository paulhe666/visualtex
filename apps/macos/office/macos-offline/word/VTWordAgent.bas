Attribute VB_Name = "VTWordAgent"
Option Explicit
Option Private Module

' ---------------------------------------------------------------------------------------------------
' AI agents (docs/office-mcp/MACOS.md): VisualTeX_AgentRun
'
' The app writes OfficeSessions/agent-request.txt (key=value rows: id, op, s.<name> = Base64URL UTF-8
' text, n.<name> = number, b.<name> = 0/1, c.<name> = array length, nested names joined by ".") and runs
' this macro through AppleScript. The answer goes to OfficeSessions/agent-response.txt:
' {"id":...,"ok":true,"data":...} or {"id":...,"ok":false,"error":{"code":...,"message":...}}.
' Reads report what the document holds (a VisualTeX formula's encoded metadata, a native equation's Word
' XML); the app turns them into LaTeX with the readers Windows shares.
' ---------------------------------------------------------------------------------------------------

Private Const VT_AGENT_MAX_REQUEST_CHARACTERS As Long = 2097152
Private Const VT_AGENT_RESPONSE_CHUNK As Long = 40000

Private Function VTAgentRequestPath() As String
    VTAgentRequestPath = VTSessionRoot() & "/agent-request.txt"
End Function

Private Function VTAgentResponsePath() As String
    VTAgentResponsePath = VTSessionRoot() & "/agent-response.txt"
End Function

Public Function VTAgentReadRequest() As Collection
    Dim values As Collection
    Dim contents As String
    Dim rows() As String
    Dim row As Variant
    Dim separator As Long

    Set values = New Collection
    contents = Replace$(VTReadText(VTAgentRequestPath(), VT_AGENT_MAX_REQUEST_CHARACTERS), vbCr, "")
    rows = Split(contents, vbLf)
    For Each row In rows
        separator = InStr(1, CStr(row), "=", vbBinaryCompare)
        If separator > 1 Then
            If Not VTCollectionHasKey(values, Left$(CStr(row), separator - 1)) Then
                values.Add Mid$(CStr(row), separator + 1), Left$(CStr(row), separator - 1)
            End If
        End If
    Next row
    If Not VTCollectionHasKey(values, "id") Or Not VTCollectionHasKey(values, "op") Then
        Err.Raise vbObjectError + 7901, "VisualTeX", "invalid_arguments|The agent request is incomplete."
    End If
    Set VTAgentReadRequest = values
End Function

Public Function VTAgentRaw(ByVal key As String) As String
    If VTCollectionHasKey(VT_AGENT_REQUEST, key) Then VTAgentRaw = CStr(VT_AGENT_REQUEST(key))
End Function

Private Function VTAgentHas(ByVal name As String) As Boolean
    VTAgentHas = VTCollectionHasKey(VT_AGENT_REQUEST, "s." & name) Or _
        VTCollectionHasKey(VT_AGENT_REQUEST, "n." & name) Or _
        VTCollectionHasKey(VT_AGENT_REQUEST, "b." & name) Or _
        VTCollectionHasKey(VT_AGENT_REQUEST, "c." & name)
End Function

Private Function VTAgentString(ByVal name As String) As String
    If VTCollectionHasKey(VT_AGENT_REQUEST, "s." & name) Then
        VTAgentString = VTBase64UrlDecodeUtf8(CStr(VT_AGENT_REQUEST("s." & name)))
    End If
End Function

Private Function VTAgentNumber(ByVal name As String, ByVal fallback As Double) As Double
    If VTCollectionHasKey(VT_AGENT_REQUEST, "n." & name) Then
        VTAgentNumber = VTParseInvariantDouble(CStr(VT_AGENT_REQUEST("n." & name)))
    Else
        VTAgentNumber = fallback
    End If
End Function

Private Function VTAgentBoolean(ByVal name As String, ByVal fallback As Boolean) As Boolean
    If VTCollectionHasKey(VT_AGENT_REQUEST, "b." & name) Then
        VTAgentBoolean = (CStr(VT_AGENT_REQUEST("b." & name)) = "1")
    Else
        VTAgentBoolean = fallback
    End If
End Function

Private Function VTAgentCount(ByVal name As String) As Long
    If VTCollectionHasKey(VT_AGENT_REQUEST, "c." & name) Then VTAgentCount = CLng(VT_AGENT_REQUEST("c." & name))
End Function

' The answer, in pieces small enough for the AppleScriptTask file bridge.
Public Sub VTAgentWriteResponse(ByVal contents As String)
    Dim position As Long
    Dim path As String

    path = VTAgentResponsePath()
    VTWriteTextAtomic path, Left$(contents, VT_AGENT_RESPONSE_CHUNK)
    position = VT_AGENT_RESPONSE_CHUNK + 1
    Do While position <= Len(contents)
        VTAppendText path, Mid$(contents, position, VT_AGENT_RESPONSE_CHUNK)
        position = position + VT_AGENT_RESPONSE_CHUNK
    Loop
End Sub

' Base64URL needs no JSON escaping (VTJsonString walks every character).
Private Function VTAgentQuotedBase64(ByVal value As String) As String
    If Len(value) = 0 Then
        VTAgentQuotedBase64 = "null"
    Else
        VTAgentQuotedBase64 = """" & VTBase64UrlEncodeUtf8(value) & """"
    End If
End Function

Private Function VTAgentDocument() As Document
    Dim key As String
    Dim found As Document

    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7902, "VisualTeX", "document_not_found|No document is open in Word."
    End If
    key = VTAgentString("document")
    If Len(key) = 0 Then
        Set VTAgentDocument = ActiveDocument
        Exit Function
    End If
    Set found = VTFindOpenWordDocumentByIdentity(key)
    If found Is Nothing Then
        Err.Raise vbObjectError + 7902, "VisualTeX", "document_not_found|The document is no longer open; call list_documents again."
    End If
    Set VTAgentDocument = found
End Function

Public Function VTAgentListDocuments() As String
    Dim result As String
    Dim candidate As Document
    Dim activeIdentity As String
    Dim identity As String

    If Documents.Count > 0 Then activeIdentity = VTWordDocumentIdentityForDocument(ActiveDocument)
    For Each candidate In Documents
        identity = VTWordDocumentIdentityForDocument(candidate)
        If Len(result) > 0 Then result = result & ","
        result = result & "{""key"":" & VTJsonString(identity) & _
            ",""name"":" & VTJsonString(candidate.Name) & _
            ",""path"":" & VTJsonNullableString(candidate.FullName) & _
            ",""active"":" & VTJsonBoolean(identity = activeIdentity) & _
            ",""saved"":" & VTJsonBoolean(candidate.Saved) & _
            ",""read_only"":" & VTJsonBoolean(candidate.ReadOnly) & "}"
    Next candidate
    VTAgentListDocuments = "{""documents"":[" & result & "]}"
End Function

' One formula of the document as the app reads it: kind, place, display, numbering, size and the payload
' (encoded metadata for a VisualTeX formula, the document part of WordOpenXML for a native equation).
Private Function VTAgentFormulaJson( _
    ByVal documentObject As Document, _
    ByVal formulaObject As Object, _
    ByVal isPicture As Boolean, _
    ByVal contextChars As Long, _
    ByVal withPayload As Boolean) As String

    Dim sourceRange As Range
    Dim formulaShape As InlineShape
    Dim formulaMath As OMath
    Dim nativeBookmark As Bookmark
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean
    Dim fontSizePt As Double
    Dim payload As String
    Dim kind As String
    Dim agentId As String
    Dim before As String
    Dim after As String
    Dim contextStart As Long
    Dim contextEnd As Long

    If isPicture Then
        Set formulaShape = formulaObject
        Set sourceRange = VTVisualTeXImageContainerRange(formulaShape)
        If sourceRange Is Nothing Then Set sourceRange = formulaShape.Range.Duplicate
        Call VTTryResolveVisualTeXInlineShapeReference( _
            formulaShape, formulaId, displayMode, numbered, encodedMetadata, _
            metadataNeedsWrite, formatNeedsWrite, False)
        fontSizePt = VTWordImageFormulaFontSize(formulaShape)
        If Not VTValidWordFormulaFontSize(fontSizePt) Then fontSizePt = VTPreferredWordFormulaFontSize(formulaShape.Range.Duplicate)
        kind = "picture"
        agentId = "vtx:" & formulaId
        payload = encodedMetadata
    Else
        Set formulaMath = formulaObject
        Set sourceRange = formulaMath.Range.Duplicate
        If VTTryFindNativeFormulaBookmarkLocally(sourceRange, nativeBookmark) Then
            If VTTryFormulaIdFromNativeBookmark(nativeBookmark.Name, formulaId) Then
                If Not VTTryReadWordFormulaFormat(documentObject, formulaId, displayMode, numbered) Then
                    displayMode = ""
                    numbered = False
                End If
            End If
        End If
        If Not VTPureNativeNumberField(sourceRange) Is Nothing Then numbered = True
        If Len(displayMode) = 0 Then
            If formulaMath.Type = wdOMathDisplay Then displayMode = "block" Else displayMode = "inline"
        End If
        fontSizePt = sourceRange.Font.Size
        If Not VTValidWordFormulaFontSize(fontSizePt) Then fontSizePt = VTPreferredWordFormulaFontSize(sourceRange)
        kind = "office_equation"
        agentId = "w:m:" & CStr(sourceRange.Start) & ":" & CStr(sourceRange.End)
        If withPayload Then payload = VTFormulaRestoreNativeXml(sourceRange)
    End If

    If contextChars > 0 Then
        contextStart = sourceRange.Start - contextChars
        If contextStart < 0 Then contextStart = 0
        before = documentObject.Range(contextStart, sourceRange.Start).Text
        contextEnd = sourceRange.End + contextChars
        If contextEnd > documentObject.Content.End Then contextEnd = documentObject.Content.End
        after = documentObject.Range(sourceRange.End, contextEnd).Text
    End If

    VTAgentFormulaJson = "{""formula_id"":" & VTJsonString(agentId) & _
        ",""kind"":" & VTJsonString(kind) & _
        ",""display"":" & VTJsonString(displayMode) & _
        ",""numbered"":" & VTJsonBoolean(numbered) & _
        ",""font_size_pt"":" & VTJsonNumber(fontSizePt) & _
        ",""start"":" & CStr(sourceRange.Start) & _
        ",""end"":" & CStr(sourceRange.End) & _
        ",""vtx_formula_id"":" & VTJsonNullableString(formulaId) & _
        ",""payload_b64"":" & VTAgentQuotedBase64(payload) & _
        ",""context_before"":" & VTJsonString(Replace$(before, vbCr, vbLf)) & _
        ",""context_after"":" & VTJsonString(Replace$(after, vbCr, vbLf)) & "}"
End Function

' The VisualTeX formulas and native equations of the main story, in reading order: pairs (start, object).
Private Function VTAgentCollectFormulas(ByVal documentObject As Document) As Collection
    Dim result As Collection
    Dim starts As Collection
    Dim formulaShape As InlineShape
    Dim formulaMath As OMath
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean
    Dim entry As Variant
    Dim index As Long
    Dim position As Long

    Set result = New Collection
    For Each formulaShape In documentObject.Content.InlineShapes
        formulaId = ""
        If VTTryResolveVisualTeXInlineShapeReference( _
           formulaShape, formulaId, displayMode, numbered, encodedMetadata, _
           metadataNeedsWrite, formatNeedsWrite, False) Then
            result.Add Array(formulaShape.Range.Start, formulaShape, True)
        End If
    Next formulaShape
    For Each formulaMath In VTOMathsInDocumentOrder(documentObject.Content)
        ' an image converted from an equation can leave an empty math zone behind
        If formulaMath.Range.InlineShapes.Count = 0 And Not VTIsEquationReferenceMath(formulaMath) And _
           Len(Trim$(Replace$(formulaMath.Range.Text, vbCr, ""))) > 0 Then
            entry = Array(formulaMath.Range.Start, formulaMath, False)
            position = result.Count + 1
            For index = 1 To result.Count
                If result(index)(0) > formulaMath.Range.Start Then
                    position = index
                    Exit For
                End If
            Next index
            If position > result.Count Then result.Add entry Else result.Add entry, Before:=position
        End If
    Next formulaMath
    Set VTAgentCollectFormulas = result
End Function

Private Function VTAgentKindWanted(ByVal kind As String) As Boolean
    Dim count As Long
    Dim index As Long

    count = VTAgentCount("kinds")
    If count = 0 Then
        VTAgentKindWanted = True
        Exit Function
    End If
    For index = 0 To count - 1
        If VTAgentString("kinds." & CStr(index)) = kind Then
            VTAgentKindWanted = True
            Exit Function
        End If
    Next index
End Function

Public Function VTAgentListFormulas() As String
    Dim documentObject As Document
    Dim formulas As Collection
    Dim offset As Long
    Dim limit As Long
    Dim contextChars As Long
    Dim index As Long
    Dim listed As Long
    Dim total As Long
    Dim items As String
    Dim revision As String
    Dim kind As String

    Set documentObject = VTAgentDocument()
    offset = CLng(VTAgentNumber("offset", 0))
    limit = CLng(VTAgentNumber("limit", 50))
    contextChars = CLng(VTAgentNumber("contextChars", 40))
    Set formulas = VTAgentCollectFormulas(documentObject)
    For index = 1 To formulas.Count
        If formulas(index)(2) Then kind = "picture" Else kind = "office_equation"
        If VTAgentKindWanted(kind) Then
            revision = revision & CStr(formulas(index)(0)) & ";"
            If total >= offset And listed < limit Then
                If Len(items) > 0 Then items = items & ","
                items = items & VTAgentFormulaJson(documentObject, formulas(index)(1), formulas(index)(2), contextChars, True)
                listed = listed + 1
            End If
            total = total + 1
        End If
    Next index
    VTAgentListFormulas = "{""revision"":" & VTJsonString(revision) & ",""total"":" & CStr(total) & _
        ",""offset"":" & CStr(offset) & ",""formulas"":[" & items & "]}"
End Function

' The formula an agent id names: "vtx:<id>" (a VisualTeX formula) or "w:m:<start>:<end>" (a native equation).
Private Function VTAgentFindFormula( _
    ByVal documentObject As Document, _
    ByVal agentId As String, _
    ByRef isPicture As Boolean) As Object

    Dim formulaShape As InlineShape
    Dim formulaMath As OMath
    Dim parts() As String
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean
    Dim startPosition As Long
    Dim endPosition As Long

    If Left$(agentId, 4) = "vtx:" Then
        For Each formulaShape In documentObject.Content.InlineShapes
            formulaId = ""
            If VTTryResolveVisualTeXInlineShapeReference( _
               formulaShape, formulaId, displayMode, numbered, encodedMetadata, _
               metadataNeedsWrite, formatNeedsWrite, False) Then
                If formulaId = Mid$(agentId, 5) Then
                    isPicture = True
                    Set VTAgentFindFormula = formulaShape
                    Exit Function
                End If
            End If
        Next formulaShape
        Err.Raise vbObjectError + 7903, "VisualTeX", "formula_not_found|No VisualTeX formula in this document has that id."
    End If
    parts = Split(agentId, ":")
    If UBound(parts) < 3 Or parts(0) <> "w" Or parts(1) <> "m" Then
        Err.Raise vbObjectError + 7903, "VisualTeX", "formula_not_found|This is not a Word formula id from list_formulas."
    End If
    startPosition = CLng(parts(2))
    endPosition = CLng(parts(3))
    If startPosition < 0 Or endPosition > documentObject.Content.End Or endPosition <= startPosition Then
        Err.Raise vbObjectError + 7904, "VisualTeX", "stale_formula_id|The document changed; call list_formulas again."
    End If
    For Each formulaMath In documentObject.Range(startPosition, endPosition).OMaths
        If formulaMath.Range.Start = startPosition And formulaMath.Range.End = endPosition Then
            isPicture = False
            Set VTAgentFindFormula = formulaMath
            Exit Function
        End If
    Next formulaMath
    Err.Raise vbObjectError + 7904, "VisualTeX", "stale_formula_id|The document changed; call list_formulas again."
End Function

Public Function VTAgentReadFormula() As String
    Dim documentObject As Document
    Dim formulaObject As Object
    Dim isPicture As Boolean

    Set documentObject = VTAgentDocument()
    Set formulaObject = VTAgentFindFormula(documentObject, VTAgentString("formulaId"), isPicture)
    VTAgentReadFormula = VTAgentFormulaJson(documentObject, formulaObject, isPicture, 120, True)
End Function

Public Function VTAgentGetSelection() As String
    Dim documentObject As Document
    Dim selected As Range
    Dim formulaShape As InlineShape
    Dim formulaMath As OMath
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean
    Dim kind As String
    Dim formula As String
    Dim preview As String

    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7902, "VisualTeX", "document_not_found|No document is open in Word."
    End If
    Set documentObject = ActiveDocument
    If Len(VTAgentString("document")) > 0 Then
        If VTWordDocumentIdentityForDocument(VTAgentDocument()) <> VTWordDocumentIdentityForDocument(ActiveDocument) Then
            ' the document's own window holds its selection
            Set documentObject = VTAgentDocument()
        End If
    End If
    Set selected = documentObject.ActiveWindow.Selection.Range
    kind = "none"
    If selected.InlineShapes.Count = 1 Then
        Set formulaShape = selected.InlineShapes(1)
        If VTTryResolveVisualTeXInlineShapeReference( _
           formulaShape, formulaId, displayMode, numbered, encodedMetadata, _
           metadataNeedsWrite, formatNeedsWrite, False) Then
            kind = "formula"
            formula = VTAgentFormulaJson(documentObject, formulaShape, True, 60, True)
        End If
    End If
    If kind = "none" And selected.OMaths.Count >= 1 Then
        Set formulaMath = selected.OMaths(1)
        If formulaMath.Range.InlineShapes.Count = 0 Then
            kind = "formula"
            formula = VTAgentFormulaJson(documentObject, formulaMath, False, 60, True)
        End If
    End If
    If kind = "none" And selected.End > selected.Start Then
        kind = "text"
        preview = selected.Text
        If Len(preview) > 200 Then preview = Left$(preview, 200) & ChrW(8230)
    End If
    If Len(formula) = 0 Then formula = "null"
    VTAgentGetSelection = "{""document"":" & VTJsonString(VTWordDocumentIdentityForDocument(documentObject)) & _
        ",""kind"":" & VTJsonString(kind) & ",""formula"":" & formula & _
        ",""text_preview"":" & VTJsonNullableString(Replace$(preview, vbCr, vbLf)) & _
        ",""location"":{""start"":" & CStr(selected.Start) & ",""end"":" & CStr(selected.End) & "}" & _
        ",""read_only"":" & VTJsonBoolean(documentObject.ReadOnly) & "}"
End Function

' ---------------------------------------------------------------------------------------------------
' AI agents: writes. Each puts its target in place (the document active, the caret or the formula selected)
' and runs the Ribbon's command with VT_AGENT_MODE on, as the Windows add-in does: errors and notices come back
' instead of dialogs. A command that opens the formula editor hands its session to the app instead
' (VTAgentSessionJson), which draws the agent's LaTeX in its hidden editor and applies it like Insert.
' ---------------------------------------------------------------------------------------------------

Private Function VTAgentWritableDocument() As Document
    Dim documentObject As Document

    Set documentObject = VTAgentDocument()
    If documentObject.ReadOnly Then
        Err.Raise vbObjectError + 7910, "VisualTeX", "document_read_only|The document is read-only."
    End If
    If VTWordDocumentIdentityForDocument(ActiveDocument) <> VTWordDocumentIdentityForDocument(documentObject) Then
        documentObject.Activate
    End If
    Set VTAgentWritableDocument = documentObject
End Function

Private Function VTAgentFormulaRange(ByVal formulaObject As Object, ByVal isPicture As Boolean) As Range
    Dim formulaShape As InlineShape
    Dim formulaMath As OMath

    If isPicture Then
        Set formulaShape = formulaObject
        Set VTAgentFormulaRange = VTVisualTeXImageContainerRange(formulaShape)
        If VTAgentFormulaRange Is Nothing Then Set VTAgentFormulaRange = formulaShape.Range.Duplicate
    Else
        Set formulaMath = formulaObject
        Set VTAgentFormulaRange = formulaMath.Range.Duplicate
    End If
End Function

Private Function VTAgentIsDisplay( _
    ByVal documentObject As Document, _
    ByVal formulaObject As Object, _
    ByVal isPicture As Boolean) As Boolean

    VTAgentIsDisplay = InStr(1, VTAgentFormulaJson(documentObject, formulaObject, isPicture, 0, False), _
        """display"":""block""", vbBinaryCompare) > 0
End Function

' A caret at a native equation's edge is still in its math zone, where a new formula would join it.
Private Sub VTAgentLeaveMathZone(ByVal caret As Range, ByVal forward As Boolean)
    Dim attempt As Long

    For attempt = 1 To 2
        If caret.OMaths.Count = 0 Then Exit Sub
        If forward Then caret.Move wdCharacter, 1 Else caret.Move wdCharacter, -1
    Next attempt
End Sub

' location: "cursor" (default), "document_end", {after_formula}, {before_formula} or {paragraph_end}.
Private Sub VTAgentPlaceInsertion(ByVal documentObject As Document)
    Dim target As Range
    Dim formulaObject As Object
    Dim isPicture As Boolean
    Dim where As String
    Dim index As Long

    where = VTAgentString("location")
    If Len(VTAgentString("location.after_formula")) > 0 Then
        Set formulaObject = VTAgentFindFormula(documentObject, VTAgentString("location.after_formula"), isPicture)
        Set target = VTAgentFormulaRange(formulaObject, isPicture)
        If VTAgentIsDisplay(documentObject, formulaObject, isPicture) Then
            Set target = target.Paragraphs(1).Range.Duplicate
            target.MoveEnd wdCharacter, -1
        End If
        target.Collapse wdCollapseEnd
        VTAgentLeaveMathZone target, True
    ElseIf Len(VTAgentString("location.before_formula")) > 0 Then
        Set formulaObject = VTAgentFindFormula(documentObject, VTAgentString("location.before_formula"), isPicture)
        Set target = VTAgentFormulaRange(formulaObject, isPicture)
        If VTAgentIsDisplay(documentObject, formulaObject, isPicture) Then Set target = target.Paragraphs(1).Range.Duplicate
        target.Collapse wdCollapseStart
        VTAgentLeaveMathZone target, False
    ElseIf VTAgentHas("location.paragraph_end") Then
        index = CLng(VTAgentNumber("location.paragraph_end", 0))
        If index < 1 Or index > documentObject.Paragraphs.Count Then
            Err.Raise vbObjectError + 7911, "VisualTeX", "invalid_arguments|The document has " & _
                CStr(documentObject.Paragraphs.Count) & " paragraphs."
        End If
        Set target = documentObject.Paragraphs(index).Range.Duplicate
        target.MoveEnd wdCharacter, -1
        target.Collapse wdCollapseEnd
    ElseIf where = "document_end" Then
        Set target = documentObject.Content.Duplicate
        target.MoveEnd wdCharacter, -1
        target.Collapse wdCollapseEnd
    ElseIf Len(where) > 0 And where <> "cursor" Then
        Err.Raise vbObjectError + 7911, "VisualTeX", "invalid_arguments|Unknown location " & where & "."
    ElseIf Selection.Start <> Selection.End Then
        ' a new formula goes after the user's selection, never over it
        Set target = Selection.Range.Duplicate
        target.Collapse wdCollapseEnd
    End If
    If Not target Is Nothing Then target.Select
End Sub

Private Function VTAgentSelectFormula( _
    ByVal documentObject As Document, _
    ByVal agentId As String, _
    ByRef isPicture As Boolean) As Object

    Dim formulaObject As Object
    Dim formulaShape As InlineShape
    Dim formulaMath As OMath

    Set formulaObject = VTAgentFindFormula(documentObject, agentId, isPicture)
    If isPicture Then
        Set formulaShape = formulaObject
        formulaShape.Select
    Else
        Set formulaMath = formulaObject
        formulaMath.Range.Select
    End If
    Set VTAgentSelectFormula = formulaObject
End Function

' What a command that opens the formula editor left: its session for the app, or nothing when the command
' finished by itself (a conversion read from the formula cache). Errors it reported become the agent's error.
Private Function VTAgentSessionJson() As String
    If Len(VT_AGENT_ERRORS) > 0 Then
        Err.Raise vbObjectError + 7912, "VisualTeX", "operation_failed|Word could not carry out the operation."
    End If
    VTAgentSessionJson = "{""session_id"":" & VTJsonNullableString(VT_AGENT_SESSION_ID) & _
        ",""request"":" & IIf(Len(VT_AGENT_REQUEST_JSON) > 0, VT_AGENT_REQUEST_JSON, "null") & _
        ",""message"":" & VTJsonNullableString(VT_AGENT_NOTES) & "}"
End Function

Private Function VTAgentDone(ByVal message As String) As String
    If Len(VT_AGENT_ERRORS) > 0 Then
        Err.Raise vbObjectError + 7912, "VisualTeX", "operation_failed|Word could not carry out the operation."
    End If
    If Len(VT_AGENT_NOTES) > 0 Then
        If Len(message) > 0 Then message = message & " "
        message = message & Replace$(VT_AGENT_NOTES, vbLf, " ")
    End If
    VTAgentDone = "{""message"":" & VTJsonNullableString(message) & "}"
End Function

Public Function VTAgentBeginInsert() As String
    Dim documentObject As Document
    Dim displayMode As String

    Set documentObject = VTAgentWritableDocument()
    VTAgentPlaceInsertion documentObject
    If VTAgentString("display") = "block" Then displayMode = "block" Else displayMode = "inline"
    VTWordCreate displayMode, VTAgentBoolean("numbered", False) And displayMode = "block", _
        VTAgentString("format") = "office_equation"
    VTAgentBeginInsert = VTAgentSessionJson()
End Function

' New LaTeX and/or another format: the editor's session for the formula, or the Ribbon's conversion of it.
Public Function VTAgentBeginUpdate() As String
    Dim documentObject As Document
    Dim isPicture As Boolean
    Dim targetFormat As String

    Set documentObject = VTAgentWritableDocument()
    Call VTAgentSelectFormula(documentObject, VTAgentString("formulaId"), isPicture)
    targetFormat = VTAgentString("format")
    If targetFormat = "office_equation" And isPicture Then
        VisualTeX_ConvertSelectedToNativeEquation
    ElseIf targetFormat = "picture" And Not isPicture Then
        VisualTeX_ConvertSelectedToImageFormula
    Else
        VisualTeX_EditSelected
    End If
    VTAgentBeginUpdate = VTAgentSessionJson()
End Function

' A VisualTeX formula by its own id (a picture's, or the anchor of an equation VisualTeX wrote).
Public Function VTAgentFindVisualTeXFormula() As String
    Dim documentObject As Document
    Dim entry As Variant
    Dim formulaShape As InlineShape
    Dim formulaMath As OMath
    Dim nativeBookmark As Bookmark
    Dim wanted As String
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim encodedMetadata As String
    Dim metadataNeedsWrite As Boolean
    Dim formatNeedsWrite As Boolean

    Set documentObject = VTAgentDocument()
    wanted = VTAgentString("vtxId")
    For Each entry In VTAgentCollectFormulas(documentObject)
        formulaId = ""
        If entry(2) Then
            Set formulaShape = entry(1)
            Call VTTryResolveVisualTeXInlineShapeReference(formulaShape, formulaId, displayMode, numbered, _
                encodedMetadata, metadataNeedsWrite, formatNeedsWrite, False)
        Else
            Set formulaMath = entry(1)
            If VTTryFindNativeFormulaBookmarkLocally(formulaMath.Range.Duplicate, nativeBookmark) Then
                Call VTTryFormulaIdFromNativeBookmark(nativeBookmark.Name, formulaId)
            End If
        End If
        If Len(formulaId) > 0 And formulaId = wanted Then
            VTAgentFindVisualTeXFormula = "{""formula"":" & _
                VTAgentFormulaJson(documentObject, entry(1), CBool(entry(2)), 120, True) & "}"
            Exit Function
        End If
    Next entry
    VTAgentFindVisualTeXFormula = "{""formula"":null}"
End Function

Public Function VTAgentSelectFormulaOperation() As String
    Dim documentObject As Document
    Dim formulaObject As Object
    Dim isPicture As Boolean

    Set documentObject = VTAgentDocument()
    If VTWordDocumentIdentityForDocument(ActiveDocument) <> VTWordDocumentIdentityForDocument(documentObject) Then
        documentObject.Activate
    End If
    Set formulaObject = VTAgentSelectFormula(documentObject, VTAgentString("formulaId"), isPicture)
    On Error Resume Next
    ActiveWindow.ScrollIntoView Selection.Range, True
    On Error GoTo 0
    VTAgentSelectFormulaOperation = "{""formula"":" & VTAgentFormulaJson(documentObject, formulaObject, isPicture, 0, True) & "}"
End Function

' A display formula goes with its paragraph (and its number); the numbers after it follow.
Public Function VTAgentDeleteFormula() As String
    Dim documentObject As Document
    Dim formulaObject As Object
    Dim isPicture As Boolean
    Dim target As Range
    Dim numbered As Boolean
    Dim recordStarted As Boolean

    Set documentObject = VTAgentWritableDocument()
    Set formulaObject = VTAgentFindFormula(documentObject, VTAgentString("formulaId"), isPicture)
    numbered = InStr(1, VTAgentFormulaJson(documentObject, formulaObject, isPicture, 0, False), _
        """numbered"":true", vbBinaryCompare) > 0
    Set target = VTAgentFormulaRange(formulaObject, isPicture)
    If VTAgentIsDisplay(documentObject, formulaObject, isPicture) Then Set target = target.Paragraphs(1).Range.Duplicate
    On Error GoTo Failed
    Application.UndoRecord.StartCustomRecord "VisualTeX Delete Formula"
    recordStarted = True
    target.Delete
    If numbered Then Call VTUpdateEquationNumbersCore(documentObject, True)
    Application.UndoRecord.EndCustomRecord
    VTAgentDeleteFormula = VTAgentDone("Deleted the formula.")
    Exit Function
Failed:
    Dim errorDescription As String
    errorDescription = Err.Description
    On Error Resume Next
    If recordStarted Then Application.UndoRecord.EndCustomRecord
    On Error GoTo 0
    Err.Raise vbObjectError + 7913, "VisualTeX", "operation_failed|" & errorDescription
End Function

Public Function VTAgentSetFontSize() As String
    Dim documentObject As Document
    Dim formulaObject As Object
    Dim formulaMath As OMath
    Dim isPicture As Boolean
    Dim size As Double

    size = VTAgentNumber("font_size_pt", 0)
    If size < 5 Or size > 200 Then
        Err.Raise vbObjectError + 7911, "VisualTeX", "invalid_arguments|font_size_pt must be from 5 to 200."
    End If
    Set documentObject = VTAgentWritableDocument()
    If Len(VTAgentString("kind")) > 0 Then
        ' every inline or every display formula, as the Ribbon's conversion sizes do
        VTAgentSetFontSize = VTAgentDone("Resized " & CStr(VTApplyFontSizeToDocumentFormulas( _
            IIf(VTAgentString("kind") = "display", "block", "inline"), size)) & " " & _
            VTAgentString("kind") & " formula(s) to " & Trim$(Str$(size)) & " pt.")
        Exit Function
    End If
    Set formulaObject = VTAgentSelectFormula(documentObject, VTAgentString("formulaId"), isPicture)
    If isPicture Then
        VisualTeX_SetSelectedImageFormulaFontSize Trim$(Str$(size))
    Else
        Set formulaMath = formulaObject
        formulaMath.Range.Font.Size = size
    End If
    VTAgentSetFontSize = VTAgentDone("The formula is " & Trim$(Str$(size)) & " pt.")
End Function

Public Function VTAgentNumberDisplayFormulas() As String
    Call VTAgentWritableDocument
    If VTAgentString("scope") = "selection" Then
        VisualTeX_NumberDisplayFormulasSelection
    Else
        VisualTeX_NumberDisplayFormulasDocument
    End If
    VTAgentNumberDisplayFormulas = VTAgentDone("")
End Function

Public Function VTAgentUpdateEquationNumbers() As String
    Dim updated As Long

    Call VTAgentWritableDocument
    updated = VTUpdateEquationNumbersCore(ActiveDocument, True)
    VTAgentUpdateEquationNumbers = VTAgentDone("Updated " & CStr(updated) & " equation numbers.")
End Function

Public Function VTAgentSetNumberFormat() As String
    Dim preset As Integer

    Select Case VTAgentString("format_id")
        Case "continuous": preset = 0
        Case "heading1-dot": preset = 1
        Case "heading1-dash": preset = 2
        Case "heading2-dot": preset = 3
        Case "heading2-dash": preset = 4
        Case Else
            Err.Raise vbObjectError + 7911, "VisualTeX", "invalid_arguments|Unknown format_id " & VTAgentString("format_id") & "."
    End Select
    Call VTAgentWritableDocument
    VTApplyEquationNumberingFormatPreset preset
    VTAgentSetNumberFormat = VTAgentDone("")
End Function

' The Ribbon's Equation-reference menu with the equation named by its number instead of picked.
Public Function VTAgentInsertReference() As String
    Dim documentObject As Document
    Dim insertionRange As Range
    Dim insertedRange As Range
    Dim items As Variant
    Dim kinds(1 To 2) As String
    Dim sources(1 To 2) As String
    Dim kindIndex As Long
    Dim itemIndex As Long
    Dim wanted As String
    Dim source As String
    Dim available As String
    Dim matches As Long
    Dim matchKind As String
    Dim matchIndex As Long

    wanted = "(" & Trim$(Replace$(Replace$(VTAgentString("number"), "(", ""), ")", "")) & ")"
    source = VTAgentString("source")
    Set documentObject = VTAgentWritableDocument()
    VTAgentPlaceInsertion documentObject
    Set insertionRange = VTResolveEquationReferenceInsertionRange(Selection.Range.Duplicate)
    If insertionRange Is Nothing Then
        Err.Raise vbObjectError + 7914, "VisualTeX", "operation_failed|A reference cannot go at this place."
    End If
    kinds(1) = VT_WORD_REFERENCE_KIND_NATIVE: sources(1) = "office_equation"
    kinds(2) = VT_WORD_REFERENCE_KIND_IMAGE: sources(2) = "picture"
    For kindIndex = 1 To 2
        items = VTEquationNumberCrossReferenceItems(documentObject, kinds(kindIndex))
        If IsArray(items) Then
            For itemIndex = LBound(items) To UBound(items)
                If Len(available) > 0 Then available = available & ", "
                available = available & CStr(items(itemIndex)) & " " & sources(kindIndex)
                If CStr(items(itemIndex)) = wanted And (Len(source) = 0 Or source = sources(kindIndex)) Then
                    matches = matches + 1
                    matchKind = kinds(kindIndex)
                    matchIndex = itemIndex
                End If
            Next itemIndex
        End If
    Next kindIndex
    If matches = 0 Then
        If Len(available) = 0 Then available = "none"
        Err.Raise vbObjectError + 7915, "VisualTeX", "formula_not_found|No numbered equation " & wanted & _
            " in this document. Numbered equations: " & available & "."
    ElseIf matches > 1 Then
        Err.Raise vbObjectError + 7915, "VisualTeX", "invalid_arguments|Several numbered equations are " & wanted & _
            "; give source to choose one. Numbered equations: " & available & "."
    End If
    Set insertedRange = VTInsertEquationNumberReferenceAtRange(insertionRange, matchIndex, True, matchKind)
    insertedRange.Collapse wdCollapseEnd
    insertedRange.Select
    VTAgentInsertReference = VTAgentDone("Inserted a reference to equation " & wanted & ".")
End Function

' Redraw, restore and conversion of many formulas: the Ribbon's commands, which open VisualTeX's progress
' window; the app waits for the session they start. A restore read from the formula cache needs none.
Public Function VTAgentDocumentOperation(ByVal operation As String) As String
    Dim scope As String
    Dim source As String
    Dim target As String

    Call VTAgentWritableDocument
    If VTAgentString("scope") = "selection" Then scope = "selection" Else scope = "document"
    Select Case operation
        Case "latex_text_to_formulas"
            If VTAgentString("format") = "office_equation" Then target = "omml" Else target = "image"
            VTStartWordLatexRedraw scope, target
        Case "formulas_to_latex_text"
            Select Case VTAgentString("from") & "|" & scope
                Case "office_equation|selection": VisualTeX_RestoreSelectionOmmlToLatex
                Case "picture|selection": VisualTeX_RestoreSelectionImageToLatex
                Case "office_equation|document": VisualTeX_RestoreDocumentOmmlToLatex
                Case "picture|document": VisualTeX_RestoreDocumentImageToLatex
                Case Else
                    Err.Raise vbObjectError + 7911, "VisualTeX", "invalid_arguments|from must be picture or office_equation."
            End Select
        Case "convert_formulas"
            Select Case VTAgentString("from") & ">" & VTAgentString("to")
                Case "picture>office_equation": source = "image": target = "omml"
                Case "office_equation>picture": source = "omml": target = "image"
                Case Else
                    Err.Raise vbObjectError + 7911, "VisualTeX", _
                        "invalid_arguments|In Word, from and to are picture and office_equation, one of each."
            End Select
            VTStartWordFormulaRestore scope, source, target
    End Select
    VTAgentDocumentOperation = VTAgentSessionJson()
End Function

' Markdown or LaTeX prose: the Ribbon's LaTeX/Markdown import at the location; the app gives the text to its
' import window.
Public Function VTAgentBeginImport() As String
    Call VTAgentPlaceInsertion(VTAgentWritableDocument())
    VisualTeX_InsertLatexMarkdownDocument
    VTAgentBeginImport = VTAgentSessionJson()
End Function
