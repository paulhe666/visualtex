Attribute VB_Name = "VTPowerPointAdapter"
Option Explicit

Private Const VT_POWERPOINT_HOST As String = "powerpoint"
Private Const VT_POWERPOINT_STATUS_FILE As String = "/OfficePluginStatus/powerpoint.json"
Private Const VT_POWERPOINT_SOURCE_REVISION As String = _
    "powerpoint-native-omml-edit-20260926-r1"
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
' AI agents: the request VisualTeX_AgentRun is answering
Private VT_PPT_AGENT_REQUEST As Collection
Private Const VT_PPT_AGENT_MAX_REQUEST_CHARACTERS As Long = 2097152
Private Const VT_PPT_AGENT_RESPONSE_CHUNK As Long = 40000

' ---------------------------------------------------------------------------------------------------
' AI agents (docs/office-mcp/MACOS.md): VisualTeX_AgentRun for PowerPoint
'
' The app writes OfficeSessions/agent-request.txt (key=value rows: id, op, s.<name> = Base64URL UTF-8
' text, n.<name> = number, b.<name> = 0/1, c.<name> = array length) and runs this macro through AppleScript;
' the answer goes to OfficeSessions/agent-response.txt as {"id","ok","data"} or {"id","ok":false,"error"}.
' A VisualTeX formula reports its encoded metadata. The object model has no OMML: the app reads a native equation
' by having VisualTeX_AgentCopyEquation copy it, with the user's clipboard saved around the copies, as a restore does.
' ---------------------------------------------------------------------------------------------------

Public Sub VisualTeX_AgentRun()
    Dim requestId As String
    Dim operation As String
    Dim data As String
    Dim code As String
    Dim message As String
    Dim separator As Long

    On Error GoTo AgentFailed
    VT_AGENT_ERRORS = ""
    VT_AGENT_NOTES = ""
    VT_AGENT_SESSION_ID = ""
    VT_AGENT_REQUEST_JSON = ""
    Set VT_PPT_AGENT_REQUEST = VTPptAgentReadRequest()
    requestId = VTPptAgentRaw("id")
    operation = VTPptAgentRaw("op")
    VT_AGENT_MODE = True
    Select Case operation
        Case "list_documents": data = VTPptAgentListDocuments()
        Case "get_selection": data = VTPptAgentGetSelection()
        Case "list_formulas": data = VTPptAgentListFormulas()
        Case "read_formula": data = VTPptAgentReadFormula()
        Case "get_slide_layout": data = VTPptAgentSlideLayout()
        Case "find_vtx_formula": data = VTPptAgentFindVisualTeXFormula()
        Case "begin_insert": data = VTPptAgentBeginInsert()
        Case "begin_update": data = VTPptAgentBeginUpdate()
        Case "select_formula": data = VTPptAgentSelectFormula()
        Case "delete_formula": data = VTPptAgentDeleteFormula()
        Case "set_formula_font_size": data = VTPptAgentSetFontSize()
        Case "set_frame": data = VTPptAgentSetFrame()
        Case "latex_text_to_formulas", "formulas_to_latex_text": data = VTPptAgentDocumentOperation(operation)
        Case Else
            Err.Raise vbObjectError + 7900, "VisualTeX", "unknown_operation|PowerPoint has no agent operation " & operation & "."
    End Select
    VT_AGENT_MODE = False
    VTPptAgentWriteResponse "{""id"":" & VTJsonString(requestId) & ",""ok"":true,""data"":" & data & "}"
    Exit Sub

AgentFailed:
    message = Err.Description
    VT_AGENT_MODE = False
    code = "operation_failed"
    separator = InStr(1, message, "|", vbBinaryCompare)
    If separator > 1 And separator < 40 And InStr(1, Left$(message, separator - 1), " ", vbBinaryCompare) = 0 Then
        code = Left$(message, separator - 1)
        message = Mid$(message, separator + 1)
    End If
    If Len(VT_AGENT_ERRORS) > 0 Then message = message & " " & VT_AGENT_ERRORS
    On Error Resume Next
    VTPptAgentWriteResponse "{""id"":" & VTJsonString(requestId) & ",""ok"":false,""error"":{""code"":" & _
        VTJsonString(code) & ",""message"":" & VTJsonString(message) & "}}"
End Sub

' Copies one native equation (formula id "ppt:<slideId>:<shapeId>:<zone>") of the request's presentation.
Public Sub VisualTeX_AgentCopyEquation(ByVal formulaId As String)
    Dim target As Presentation
    Dim foundSlide As Slide
    Dim found As Shape
    Dim zone As Long

    ' A failure leaves the clipboard as it is, which the app reports; no dialog may open.
    On Error GoTo CopyFailed
    Set VT_PPT_AGENT_REQUEST = VTPptAgentReadRequest()
    Set target = VTPptAgentPresentation()
    Set found = VTPptAgentFindFormula(target, formulaId, foundSlide, zone)
    If zone < 1 Then Exit Sub
    found.TextFrame2.TextRange.MathZones(zone, 1).Copy
CopyFailed:
End Sub

Private Function VTPptAgentReadRequest() As Collection
    Dim values As Collection
    Dim contents As String
    Dim rows() As String
    Dim row As Variant
    Dim separator As Long

    Set values = New Collection
    contents = Replace$(VTReadText(VTSessionRoot() & "/agent-request.txt", VT_PPT_AGENT_MAX_REQUEST_CHARACTERS), vbCr, "")
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
    Set VTPptAgentReadRequest = values
End Function

Private Function VTPptAgentRaw(ByVal key As String) As String
    If VTCollectionHasKey(VT_PPT_AGENT_REQUEST, key) Then VTPptAgentRaw = CStr(VT_PPT_AGENT_REQUEST(key))
End Function

Private Function VTPptAgentString(ByVal name As String) As String
    If VTCollectionHasKey(VT_PPT_AGENT_REQUEST, "s." & name) Then
        VTPptAgentString = VTBase64UrlDecodeUtf8(CStr(VT_PPT_AGENT_REQUEST("s." & name)))
    End If
End Function

Private Function VTPptAgentNumber(ByVal name As String, ByVal fallback As Double) As Double
    If VTCollectionHasKey(VT_PPT_AGENT_REQUEST, "n." & name) Then
        VTPptAgentNumber = VTParseInvariantDouble(CStr(VT_PPT_AGENT_REQUEST("n." & name)))
    Else
        VTPptAgentNumber = fallback
    End If
End Function

Private Function VTPptAgentCount(ByVal name As String) As Long
    If VTCollectionHasKey(VT_PPT_AGENT_REQUEST, "c." & name) Then VTPptAgentCount = CLng(VT_PPT_AGENT_REQUEST("c." & name))
End Function

Private Sub VTPptAgentWriteResponse(ByVal contents As String)
    Dim position As Long
    Dim path As String

    path = VTSessionRoot() & "/agent-response.txt"
    VTWriteTextAtomic path, Left$(contents, VT_PPT_AGENT_RESPONSE_CHUNK)
    position = VT_PPT_AGENT_RESPONSE_CHUNK + 1
    Do While position <= Len(contents)
        VTAppendText path, Mid$(contents, position, VT_PPT_AGENT_RESPONSE_CHUNK)
        position = position + VT_PPT_AGENT_RESPONSE_CHUNK
    Loop
End Sub

Private Function VTPptAgentQuotedBase64(ByVal value As String) As String
    If Len(value) = 0 Then
        VTPptAgentQuotedBase64 = "null"
    Else
        VTPptAgentQuotedBase64 = """" & VTBase64UrlEncodeUtf8(value) & """"
    End If
End Function

Private Function VTPptAgentKey(ByVal target As Presentation) As String
    If Len(target.Path) > 0 Then
        VTPptAgentKey = target.FullName
    Else
        VTPptAgentKey = target.Name
    End If
End Function

Private Function VTPptAgentPresentation() As Presentation
    Dim key As String
    Dim candidate As Presentation

    If Presentations.Count = 0 Then
        Err.Raise vbObjectError + 7902, "VisualTeX", "document_not_found|No presentation is open in PowerPoint."
    End If
    key = VTPptAgentString("document")
    If Len(key) = 0 Then
        Set VTPptAgentPresentation = ActivePresentation
        Exit Function
    End If
    For Each candidate In Presentations
        If VTPptAgentKey(candidate) = key Then
            Set VTPptAgentPresentation = candidate
            Exit Function
        End If
    Next candidate
    Err.Raise vbObjectError + 7902, "VisualTeX", "document_not_found|The presentation is no longer open; call list_documents again."
End Function

Private Function VTPptAgentFrameJson(ByVal target As Shape) As String
    Dim rotation As Double
    Dim zOrder As Long

    On Error Resume Next
    rotation = target.Rotation
    zOrder = target.ZOrderPosition
    On Error GoTo 0
    VTPptAgentFrameJson = "{""left_pt"":" & VTJsonNumber(Round(CDbl(target.Left), 2)) & _
        ",""top_pt"":" & VTJsonNumber(Round(CDbl(target.Top), 2)) & _
        ",""width_pt"":" & VTJsonNumber(Round(CDbl(target.Width), 2)) & _
        ",""height_pt"":" & VTJsonNumber(Round(CDbl(target.Height), 2)) & _
        ",""rotation"":" & VTJsonNumber(Round(rotation, 2)) & _
        ",""z_order"":" & CStr(zOrder) & "}"
End Function

Private Function VTPptAgentListDocuments() As String
    Dim result As String
    Dim candidate As Presentation
    Dim activeKey As String

    On Error Resume Next
    If Presentations.Count > 0 Then activeKey = VTPptAgentKey(ActivePresentation)
    On Error GoTo 0
    For Each candidate In Presentations
        If Len(result) > 0 Then result = result & ","
        result = result & "{""key"":" & VTJsonString(VTPptAgentKey(candidate)) & _
            ",""name"":" & VTJsonString(candidate.Name) & _
            ",""path"":" & VTJsonNullableString(IIf(Len(candidate.Path) > 0, candidate.FullName, "")) & _
            ",""active"":" & VTJsonBoolean(VTPptAgentKey(candidate) = activeKey) & _
            ",""saved"":" & VTJsonBoolean(candidate.Saved <> 0) & _
            ",""read_only"":" & VTJsonBoolean(candidate.ReadOnly <> 0) & _
            ",""slide_count"":" & CStr(candidate.Slides.Count) & _
            ",""slide_width_pt"":" & VTJsonNumber(Round(CDbl(candidate.PageSetup.SlideWidth), 2)) & _
            ",""slide_height_pt"":" & VTJsonNumber(Round(CDbl(candidate.PageSetup.SlideHeight), 2)) & "}"
    Next candidate
    VTPptAgentListDocuments = "{""documents"":[" & result & "]}"
End Function

Private Function VTPptAgentMathZoneCount(ByVal target As Shape) As Long
    On Error Resume Next
    If target.HasTextFrame Then VTPptAgentMathZoneCount = target.TextFrame2.TextRange.MathZones.Count
End Function

' One formula: a VisualTeX picture (zone 0) or the zone-th native equation of a text shape.
Private Function VTPptAgentFormulaJson( _
    ByVal currentSlide As Slide, _
    ByVal target As Shape, _
    ByVal zone As Long, _
    ByVal contextChars As Long) As String

    Dim kind As String
    Dim agentId As String
    Dim payload As String
    Dim fontSizePt As Double
    Dim fullText As Object
    Dim mathZone As Object
    Dim textValue As String
    Dim zoneStart As Long
    Dim zoneLength As Long
    Dim before As String
    Dim after As String
    Dim fromIndex As Long
    Dim display As String

    display = "inline"
    If zone = 0 Then
        kind = "picture"
        agentId = "vtx:" & VTShapeFormulaId(target)
        payload = VTShapeMetadata(target)
        On Error Resume Next
        fontSizePt = VTParseInvariantDouble(target.Tags(VT_POWERPOINT_FONT_TAG))
        On Error GoTo 0
    Else
        kind = "office_equation"
        agentId = "ppt:" & CStr(currentSlide.SlideID) & ":" & CStr(target.Id) & ":" & CStr(zone)
        On Error Resume Next
        Set fullText = target.TextFrame2.TextRange
        Set mathZone = fullText.MathZones(zone, 1)
        textValue = fullText.Text
        zoneStart = mathZone.Start
        zoneLength = mathZone.Length
        fontSizePt = mathZone.Font.Size
        If Len(Trim$(Replace$(mathZone.Paragraphs(1).Text, vbCr, ""))) <= zoneLength Then display = "block"
        On Error GoTo 0
        If contextChars > 0 Then
            If zoneStart > 0 Then
                fromIndex = zoneStart - contextChars
                If fromIndex < 1 Then fromIndex = 1
                before = Mid$(textValue, fromIndex, zoneStart - fromIndex)
                after = Mid$(textValue, zoneStart + zoneLength, contextChars)
            End If
        End If
    End If

    VTPptAgentFormulaJson = "{""formula_id"":" & VTJsonString(agentId) & _
        ",""kind"":" & VTJsonString(kind) & _
        ",""display"":" & VTJsonString(display) & _
        ",""numbered"":false" & _
        ",""font_size_pt"":" & IIf(fontSizePt > 0, VTJsonNumber(Round(CDbl(fontSizePt), 2)), "null") & _
        ",""slide"":" & CStr(currentSlide.SlideIndex) & _
        ",""slide_id"":" & CStr(currentSlide.SlideID) & _
        ",""shape_id"":" & CStr(target.Id) & _
        ",""shape_name"":" & VTJsonString(target.Name) & _
        ",""equation"":" & IIf(zone > 0, CStr(zone), "null") & _
        ",""frame"":" & VTPptAgentFrameJson(target) & _
        ",""frame_of"":" & VTJsonString(IIf(zone > 0, "text_box", "formula")) & _
        ",""payload_b64"":" & VTPptAgentQuotedBase64(payload) & _
        ",""context_before"":" & VTJsonString(Replace$(Replace$(before, vbCr, vbLf), Chr$(11), vbLf)) & _
        ",""context_after"":" & VTJsonString(Replace$(Replace$(after, vbCr, vbLf), Chr$(11), vbLf)) & "}"
End Function

Private Function VTPptAgentKindWanted(ByVal kind As String) As Boolean
    Dim count As Long
    Dim index As Long

    count = VTPptAgentCount("kinds")
    If count = 0 Then
        VTPptAgentKindWanted = True
        Exit Function
    End If
    For index = 0 To count - 1
        If VTPptAgentString("kinds." & CStr(index)) = kind Then
            VTPptAgentKindWanted = True
            Exit Function
        End If
    Next index
End Function

Private Function VTPptAgentListFormulas() As String
    Dim target As Presentation
    Dim currentSlide As Slide
    Dim candidate As Shape
    Dim offset As Long
    Dim limit As Long
    Dim contextChars As Long
    Dim listed As Long
    Dim total As Long
    Dim zone As Long
    Dim zones As Long
    Dim items As String
    Dim revision As String

    Set target = VTPptAgentPresentation()
    offset = CLng(VTPptAgentNumber("offset", 0))
    limit = CLng(VTPptAgentNumber("limit", 50))
    contextChars = CLng(VTPptAgentNumber("contextChars", 40))
    For Each currentSlide In target.Slides
        For Each candidate In currentSlide.Shapes
            If VTIsVisualTeXPowerPointShape(candidate) Then
                If VTPptAgentKindWanted("picture") Then
                    revision = revision & CStr(currentSlide.SlideID) & "." & CStr(candidate.Id) & ";"
                    If total >= offset And listed < limit Then
                        If Len(items) > 0 Then items = items & ","
                        items = items & VTPptAgentFormulaJson(currentSlide, candidate, 0, contextChars)
                        listed = listed + 1
                    End If
                    total = total + 1
                End If
            ElseIf VTPptAgentKindWanted("office_equation") Then
                zones = VTPptAgentMathZoneCount(candidate)
                For zone = 1 To zones
                    revision = revision & CStr(currentSlide.SlideID) & "." & CStr(candidate.Id) & "." & CStr(zone) & ";"
                    If total >= offset And listed < limit Then
                        If Len(items) > 0 Then items = items & ","
                        items = items & VTPptAgentFormulaJson(currentSlide, candidate, zone, contextChars)
                        listed = listed + 1
                    End If
                    total = total + 1
                Next zone
            End If
        Next candidate
    Next currentSlide
    VTPptAgentListFormulas = "{""revision"":" & VTJsonString(revision) & ",""total"":" & CStr(total) & _
        ",""offset"":" & CStr(offset) & ",""formulas"":[" & items & "]}"
End Function

' The formula an agent id names: "vtx:<id>" or "ppt:<slideId>:<shapeId>:<zone>".
Private Function VTPptAgentFindFormula( _
    ByVal target As Presentation, _
    ByVal agentId As String, _
    ByRef foundSlide As Slide, _
    ByRef zone As Long) As Shape

    Dim currentSlide As Slide
    Dim candidate As Shape
    Dim parts() As String

    If Left$(agentId, 4) = "vtx:" Then
        For Each currentSlide In target.Slides
            For Each candidate In currentSlide.Shapes
                If VTIsVisualTeXPowerPointShape(candidate) Then
                    If VTShapeFormulaId(candidate) = Mid$(agentId, 5) Then
                        Set foundSlide = currentSlide
                        zone = 0
                        Set VTPptAgentFindFormula = candidate
                        Exit Function
                    End If
                End If
            Next candidate
        Next currentSlide
        Err.Raise vbObjectError + 7903, "VisualTeX", "formula_not_found|No VisualTeX formula in this presentation has that id."
    End If
    parts = Split(agentId, ":")
    If UBound(parts) <> 3 Or parts(0) <> "ppt" Then
        Err.Raise vbObjectError + 7903, "VisualTeX", "formula_not_found|This is not a PowerPoint formula id from list_formulas."
    End If
    For Each currentSlide In target.Slides
        If CStr(currentSlide.SlideID) = parts(1) Then
            For Each candidate In currentSlide.Shapes
                If CStr(candidate.Id) = parts(2) Then
                    zone = CLng(parts(3))
                    If zone < 1 Or zone > VTPptAgentMathZoneCount(candidate) Then Exit For
                    Set foundSlide = currentSlide
                    Set VTPptAgentFindFormula = candidate
                    Exit Function
                End If
            Next candidate
        End If
    Next currentSlide
    Err.Raise vbObjectError + 7904, "VisualTeX", "stale_formula_id|The presentation changed; call list_formulas again."
End Function

Private Function VTPptAgentReadFormula() As String
    Dim target As Presentation
    Dim foundSlide As Slide
    Dim found As Shape
    Dim zone As Long

    Set target = VTPptAgentPresentation()
    Set found = VTPptAgentFindFormula(target, VTPptAgentString("formulaId"), foundSlide, zone)
    VTPptAgentReadFormula = VTPptAgentFormulaJson(foundSlide, found, zone, 120)
End Function

Private Function VTPptAgentShapeType(ByVal target As Shape) As String
    Dim typeValue As Long

    On Error Resume Next
    typeValue = target.Type
    On Error GoTo 0
    Select Case typeValue
        Case 14: VTPptAgentShapeType = "placeholder"
        Case 17: VTPptAgentShapeType = "text_box"
        Case 13, 11: VTPptAgentShapeType = "picture"
        Case 6: VTPptAgentShapeType = "group"
        Case 19: VTPptAgentShapeType = "table"
        Case 3: VTPptAgentShapeType = "chart"
        Case 24: VTPptAgentShapeType = "smart_art"
        Case 9: VTPptAgentShapeType = "line"
        Case 16: VTPptAgentShapeType = "media"
        Case 7, 10, 12: VTPptAgentShapeType = "ole_object"
        Case 1, 5: VTPptAgentShapeType = "shape"
        Case Else: VTPptAgentShapeType = "other"
    End Select
End Function

Private Function VTPptAgentPlaceholderType(ByVal target As Shape) As String
    Dim typeValue As Long

    On Error GoTo NotPlaceholder
    If target.Type <> 14 Then Exit Function
    typeValue = target.PlaceholderFormat.Type
    Select Case typeValue
        Case 1, 3, 5: VTPptAgentPlaceholderType = "title"
        Case 4: VTPptAgentPlaceholderType = "subtitle"
        Case 2, 6: VTPptAgentPlaceholderType = "body"
        Case 7: VTPptAgentPlaceholderType = "content"
        Case 18: VTPptAgentPlaceholderType = "picture"
        Case 15: VTPptAgentPlaceholderType = "footer"
        Case 13: VTPptAgentPlaceholderType = "slide_number"
        Case 16: VTPptAgentPlaceholderType = "date"
        Case Else: VTPptAgentPlaceholderType = "other"
    End Select
NotPlaceholder:
End Function

Private Function VTPptAgentTextPreview(ByVal target As Shape) As String
    Dim value As String

    On Error Resume Next
    If target.HasTextFrame Then
        If target.TextFrame.HasText Then value = target.TextFrame.TextRange.Text
    End If
    On Error GoTo 0
    value = Replace$(Replace$(value, vbCr, vbLf), Chr$(11), vbLf)
    If Len(value) > 200 Then value = Left$(value, 200) & ChrW(8230)
    VTPptAgentTextPreview = value
End Function

Private Function VTPptAgentSlideLayout() As String
    Dim target As Presentation
    Dim currentSlide As Slide
    Dim candidate As Shape
    Dim slideNumber As Long
    Dim items As String
    Dim ids As String
    Dim kind As String
    Dim zone As Long
    Dim zones As Long
    Dim layoutName As String

    Set target = VTPptAgentPresentation()
    slideNumber = CLng(VTPptAgentNumber("slide", 0))
    If slideNumber = 0 Then
        On Error Resume Next
        slideNumber = ActiveWindow.View.Slide.SlideIndex
        On Error GoTo 0
        If slideNumber = 0 Then slideNumber = 1
    End If
    If slideNumber < 1 Or slideNumber > target.Slides.Count Then
        Err.Raise vbObjectError + 7905, "VisualTeX", "invalid_arguments|The presentation has " & CStr(target.Slides.Count) & " slides."
    End If
    Set currentSlide = target.Slides(slideNumber)
    For Each candidate In currentSlide.Shapes
        ids = ""
        If VTIsVisualTeXPowerPointShape(candidate) Then
            kind = "formula"
            ids = VTJsonString("vtx:" & VTShapeFormulaId(candidate))
        Else
            kind = VTPptAgentShapeType(candidate)
            zones = VTPptAgentMathZoneCount(candidate)
            For zone = 1 To zones
                If Len(ids) > 0 Then ids = ids & ","
                ids = ids & VTJsonString("ppt:" & CStr(currentSlide.SlideID) & ":" & CStr(candidate.Id) & ":" & CStr(zone))
            Next zone
        End If
        If Len(items) > 0 Then items = items & ","
        items = items & "{""shape_id"":" & CStr(candidate.Id) & _
            ",""name"":" & VTJsonString(candidate.Name) & _
            ",""type"":" & VTJsonString(kind) & _
            ",""placeholder"":" & VTJsonNullableString(VTPptAgentPlaceholderType(candidate)) & _
            ",""frame"":" & VTPptAgentFrameJson(candidate) & _
            ",""text"":" & VTJsonNullableString(VTPptAgentTextPreview(candidate)) & _
            ",""formula_ids"":" & IIf(Len(ids) > 0, "[" & ids & "]", "null") & "}"
    Next candidate
    On Error Resume Next
    layoutName = currentSlide.CustomLayout.Name
    On Error GoTo 0
    VTPptAgentSlideLayout = "{""slide"":" & CStr(slideNumber) & _
        ",""slide_id"":" & CStr(currentSlide.SlideID) & _
        ",""slide_count"":" & CStr(target.Slides.Count) & _
        ",""layout"":" & VTJsonNullableString(layoutName) & _
        ",""slide_width_pt"":" & VTJsonNumber(Round(CDbl(target.PageSetup.SlideWidth), 2)) & _
        ",""slide_height_pt"":" & VTJsonNumber(Round(CDbl(target.PageSetup.SlideHeight), 2)) & _
        ",""shapes"":[" & items & "]}"
End Function

Private Function VTPptAgentGetSelection() As String
    Dim target As Presentation
    Dim selectionType As Long
    Dim currentSlide As Slide
    Dim selectedShape As Shape
    Dim kind As String
    Dim formula As String
    Dim preview As String
    Dim slideIndex As Long
    Dim slideId As Long
    Dim shapeId As Long
    Dim zone As Long
    Dim caret As Long
    Dim fullText As Object

    Set target = VTPptAgentPresentation()
    kind = "none"
    On Error Resume Next
    Set currentSlide = ActiveWindow.View.Slide
    selectionType = ActiveWindow.Selection.Type
    On Error GoTo 0
    If Not currentSlide Is Nothing Then
        slideIndex = currentSlide.SlideIndex
        slideId = currentSlide.SlideID
    End If
    If selectionType = 1 Then kind = "slide"
    If selectionType = 2 Or selectionType = 3 Then
        On Error Resume Next
        If ActiveWindow.Selection.ShapeRange.Count = 1 Then Set selectedShape = ActiveWindow.Selection.ShapeRange(1)
        On Error GoTo 0
    End If
    If Not selectedShape Is Nothing And Not currentSlide Is Nothing Then
        shapeId = selectedShape.Id
        kind = "shape"
        If VTIsVisualTeXPowerPointShape(selectedShape) Then
            kind = "formula"
            formula = VTPptAgentFormulaJson(currentSlide, selectedShape, 0, 0)
        ElseIf VTPptAgentMathZoneCount(selectedShape) > 0 Then
            If selectionType = 3 Then
                On Error Resume Next
                caret = ActiveWindow.Selection.TextRange2.Start
                Set fullText = selectedShape.TextFrame2.TextRange
                For zone = 1 To fullText.MathZones.Count
                    If caret >= fullText.MathZones(zone, 1).Start And caret <= fullText.MathZones(zone, 1).Start + fullText.MathZones(zone, 1).Length Then Exit For
                Next zone
                On Error GoTo 0
                If zone >= 1 And zone <= VTPptAgentMathZoneCount(selectedShape) Then
                    kind = "formula"
                    formula = VTPptAgentFormulaJson(currentSlide, selectedShape, zone, 60)
                End If
            ElseIf VTPptAgentMathZoneCount(selectedShape) = 1 Then
                kind = "formula"
                formula = VTPptAgentFormulaJson(currentSlide, selectedShape, 1, 60)
            End If
        End If
        If kind = "shape" And selectionType = 3 Then
            kind = "text"
            On Error Resume Next
            preview = ActiveWindow.Selection.TextRange.Text
            On Error GoTo 0
            If Len(preview) > 200 Then preview = Left$(preview, 200) & ChrW(8230)
        End If
    End If
    If Len(formula) = 0 Then formula = "null"
    VTPptAgentGetSelection = "{""document"":" & VTJsonString(VTPptAgentKey(target)) & _
        ",""kind"":" & VTJsonString(kind) & ",""formula"":" & formula & _
        ",""text_preview"":" & VTJsonNullableString(Replace$(preview, vbCr, vbLf)) & _
        ",""location"":{""slide"":" & IIf(slideIndex > 0, CStr(slideIndex), "null") & _
        ",""slide_id"":" & IIf(slideId > 0, CStr(slideId), "null") & _
        ",""shape_id"":" & IIf(shapeId > 0, CStr(shapeId), "null") & "}" & _
        ",""read_only"":" & VTJsonBoolean(target.ReadOnly <> 0) & "}"
End Function

' ---- writes: each puts its target in place (the presentation's window and slide, the formula selected) and
' runs the Ribbon's command with VT_AGENT_MODE on; a command that opens the formula editor hands its session
' to the app (VTPptAgentSessionJson), which draws the agent's LaTeX in its hidden editor and applies it.

Private Function VTPptAgentWritablePresentation() As Presentation
    Dim target As Presentation

    Set target = VTPptAgentPresentation()
    If target.ReadOnly <> 0 Then
        Err.Raise vbObjectError + 7910, "VisualTeX", "document_read_only|The presentation is read-only."
    End If
    If VTPptAgentKey(ActivePresentation) <> VTPptAgentKey(target) Then
        If target.Windows.Count = 0 Then
            Err.Raise vbObjectError + 7910, "VisualTeX", "operation_failed|The presentation has no window."
        End If
        target.Windows(1).Activate
    End If
    Set VTPptAgentWritablePresentation = target
End Function

Private Sub VTPptAgentShowSlide(ByVal slideIndex As Long)
    If ActiveWindow.ViewType <> ppViewNormal Then ActiveWindow.ViewType = ppViewNormal
    If ActiveWindow.View.Slide.SlideIndex <> slideIndex Then ActiveWindow.View.GotoSlide slideIndex
End Sub

' The formula an agent id names, its slide shown and the formula selected (a native equation's text).
Private Function VTPptAgentSelect( _
    ByVal target As Presentation, _
    ByVal agentId As String, _
    ByRef foundSlide As Slide, _
    ByRef zone As Long) As Shape

    Dim found As Shape
    Dim mathZone As Object

    Set found = VTPptAgentFindFormula(target, agentId, foundSlide, zone)
    VTPptAgentShowSlide foundSlide.SlideIndex
    If zone = 0 Then
        found.Select
    Else
        Set mathZone = found.TextFrame2.TextRange.MathZones(zone, 1)
        found.TextFrame.TextRange.Characters(mathZone.Start, mathZone.Length).Select
    End If
    Set VTPptAgentSelect = found
End Function

Private Function VTPptAgentSessionJson() As String
    Dim slideIndex As Long

    If Len(VT_AGENT_ERRORS) > 0 Then
        Err.Raise vbObjectError + 7912, "VisualTeX", "operation_failed|PowerPoint could not carry out the operation."
    End If
    On Error Resume Next
    slideIndex = ActiveWindow.View.Slide.SlideIndex
    On Error GoTo 0
    VTPptAgentSessionJson = "{""session_id"":" & VTJsonNullableString(VT_AGENT_SESSION_ID) & _
        ",""request"":" & IIf(Len(VT_AGENT_REQUEST_JSON) > 0, VT_AGENT_REQUEST_JSON, "null") & _
        ",""slide"":" & IIf(slideIndex > 0, CStr(slideIndex), "null") & _
        ",""message"":" & VTJsonNullableString(VT_AGENT_NOTES) & "}"
End Function

Private Function VTPptAgentDone(ByVal message As String) As String
    If Len(VT_AGENT_ERRORS) > 0 Then
        Err.Raise vbObjectError + 7912, "VisualTeX", "operation_failed|PowerPoint could not carry out the operation."
    End If
    If Len(VT_AGENT_NOTES) > 0 Then
        If Len(message) > 0 Then message = message & " "
        message = message & Replace$(VT_AGENT_NOTES, vbLf, " ")
    End If
    VTPptAgentDone = "{""message"":" & VTJsonNullableString(message) & "}"
End Function

' location: "cursor" (default: the selected text box, or a new formula on the current slide) or {slide, ...}
' (a new formula on that slide; the app moves it to left_pt/top_pt afterwards).
Private Function VTPptAgentBeginInsert() As String
    Dim target As Presentation
    Dim slideNumber As Long

    Set target = VTPptAgentWritablePresentation()
    slideNumber = CLng(VTPptAgentNumber("location.slide", 0))
    If slideNumber <> 0 Then
        If slideNumber < 1 Or slideNumber > target.Slides.Count Then
            Err.Raise vbObjectError + 7911, "VisualTeX", "invalid_arguments|The presentation has " & _
                CStr(target.Slides.Count) & " slides."
        End If
        VTPptAgentShowSlide slideNumber
        ActiveWindow.Selection.Unselect
    ElseIf Len(VTPptAgentString("location")) > 0 And VTPptAgentString("location") <> "cursor" Then
        Err.Raise vbObjectError + 7911, "VisualTeX", "invalid_arguments|Unknown location " & VTPptAgentString("location") & "."
    End If
    If VTPptAgentString("format") = "office_equation" Then
        If VTPptAgentString("display") = "block" Then
            VisualTeX_NewDisplayNativeEquation
        Else
            VisualTeX_NewInlineNativeEquation
        End If
    Else
        VisualTeX_NewFormula
    End If
    VTPptAgentBeginInsert = VTPptAgentSessionJson()
End Function

Private Function VTPptAgentBeginUpdate() As String
    Dim target As Presentation
    Dim foundSlide As Slide
    Dim found As Shape
    Dim zone As Long

    Set target = VTPptAgentWritablePresentation()
    Set found = VTPptAgentSelect(target, VTPptAgentString("formulaId"), foundSlide, zone)
    If zone = 0 Then
        VisualTeX_EditShape found
    ElseIf Not VTEditSelectedNativeEquation() Then
        Err.Raise vbObjectError + 7913, "VisualTeX", "operation_failed|PowerPoint did not open the equation."
    End If
    VTPptAgentBeginUpdate = VTPptAgentSessionJson()
End Function

Private Function VTPptAgentFindVisualTeXFormula() As String
    Dim target As Presentation
    Dim currentSlide As Slide
    Dim candidate As Shape

    Set target = VTPptAgentPresentation()
    For Each currentSlide In target.Slides
        For Each candidate In currentSlide.Shapes
            If VTIsVisualTeXPowerPointShape(candidate) Then
                If VTShapeFormulaId(candidate) = VTPptAgentString("vtxId") Then
                    VTPptAgentFindVisualTeXFormula = "{""formula"":" & VTPptAgentFormulaJson(currentSlide, candidate, 0, 0) & "}"
                    Exit Function
                End If
            End If
        Next candidate
    Next currentSlide
    VTPptAgentFindVisualTeXFormula = "{""formula"":null}"
End Function

Private Function VTPptAgentSelectFormula() As String
    Dim target As Presentation
    Dim foundSlide As Slide
    Dim found As Shape
    Dim zone As Long

    Set target = VTPptAgentPresentation()
    If VTPptAgentKey(ActivePresentation) <> VTPptAgentKey(target) And target.Windows.Count > 0 Then target.Windows(1).Activate
    Set found = VTPptAgentSelect(target, VTPptAgentString("formulaId"), foundSlide, zone)
    VTPptAgentSelectFormula = "{""formula"":" & VTPptAgentFormulaJson(foundSlide, found, zone, 0) & "}"
End Function

' A native equation alone in its text box goes with the box.
Private Function VTPptAgentDeleteFormula() As String
    Dim target As Presentation
    Dim foundSlide As Slide
    Dim found As Shape
    Dim zone As Long
    Dim mathZone As Object

    Set target = VTPptAgentWritablePresentation()
    Set found = VTPptAgentFindFormula(target, VTPptAgentString("formulaId"), foundSlide, zone)
    If zone = 0 Then
        found.Delete
    Else
        Set mathZone = found.TextFrame2.TextRange.MathZones(zone, 1)
        If Len(Trim$(Replace$(Replace$(found.TextFrame.TextRange.Text, vbCr, ""), Chr$(11), ""))) <= mathZone.Length Then
            found.Delete
        Else
            mathZone.Delete
        End If
    End If
    VTPptAgentDeleteFormula = VTPptAgentDone("Deleted the formula.")
End Function

Private Function VTPptAgentSetFontSize() As String
    Dim target As Presentation
    Dim foundSlide As Slide
    Dim found As Shape
    Dim zone As Long
    Dim size As Double

    size = VTPptAgentNumber("font_size_pt", 0)
    If size < 5 Or size > 200 Then
        Err.Raise vbObjectError + 7911, "VisualTeX", "invalid_arguments|font_size_pt must be from 5 to 200."
    End If
    Set target = VTPptAgentWritablePresentation()
    Set found = VTPptAgentSelect(target, VTPptAgentString("formulaId"), foundSlide, zone)
    If zone = 0 Then
        VisualTeX_SetSelectedFormulaFontSize Trim$(Str$(size))
    Else
        found.TextFrame2.TextRange.MathZones(zone, 1).Font.Size = size
    End If
    VTPptAgentSetFontSize = VTPptAgentDone("The formula is " & Trim$(Str$(size)) & " pt.")
End Function

' position: the new frame in points; what is left out stays. width or height alone keeps the proportions
' (a text box only takes the size given).
Private Function VTPptAgentSetFrame() As String
    Dim target As Presentation
    Dim foundSlide As Slide
    Dim found As Shape
    Dim zone As Long
    Dim locked As Long
    Dim ratio As Double
    Dim hasWidth As Boolean
    Dim hasHeight As Boolean

    Set target = VTPptAgentWritablePresentation()
    Set found = VTPptAgentFindFormula(target, VTPptAgentString("formulaId"), foundSlide, zone)
    hasWidth = VTCollectionHasKey(VT_PPT_AGENT_REQUEST, "n.position.width_pt")
    hasHeight = VTCollectionHasKey(VT_PPT_AGENT_REQUEST, "n.position.height_pt")
    locked = found.LockAspectRatio
    found.LockAspectRatio = msoFalse
    If found.Height > 0 Then ratio = found.Width / found.Height
    If hasWidth And hasHeight Then
        found.Width = VTPptAgentNumber("position.width_pt", found.Width)
        found.Height = VTPptAgentNumber("position.height_pt", found.Height)
    ElseIf hasWidth Then
        found.Width = VTPptAgentNumber("position.width_pt", found.Width)
        If zone = 0 And ratio > 0 Then found.Height = found.Width / ratio
    ElseIf hasHeight Then
        found.Height = VTPptAgentNumber("position.height_pt", found.Height)
        If zone = 0 And ratio > 0 Then found.Width = found.Height * ratio
    End If
    If VTCollectionHasKey(VT_PPT_AGENT_REQUEST, "n.position.left_pt") Then found.Left = VTPptAgentNumber("position.left_pt", found.Left)
    If VTCollectionHasKey(VT_PPT_AGENT_REQUEST, "n.position.top_pt") Then found.Top = VTPptAgentNumber("position.top_pt", found.Top)
    found.LockAspectRatio = locked
    VTPptAgentSetFrame = "{""frame"":" & VTPptAgentFrameJson(found) & _
        ",""frame_of"":" & VTJsonString(IIf(zone > 0, "text_box", "formula")) & "}"
End Function

' The Ribbon's LaTeX redraw and formulas-to-LaTeX, which open VisualTeX's progress window; the app waits for
' the session they start.
Private Function VTPptAgentDocumentOperation(ByVal operation As String) As String
    Dim scope As String

    Call VTPptAgentWritablePresentation
    scope = VTPptAgentString("scope")
    If scope <> "selection" And scope <> "slide" Then scope = "presentation"
    If operation = "latex_text_to_formulas" Then
        Select Case scope & "|" & IIf(VTPptAgentString("format") = "office_equation", "omml", "image")
            Case "selection|omml": VisualTeX_RedrawSelectionToOmml
            Case "slide|omml": VisualTeX_RedrawSlideToOmml
            Case "presentation|omml": VisualTeX_RedrawPresentationToOmml
            Case "selection|image": VisualTeX_RedrawSelectionToImage
            Case "slide|image": VisualTeX_RedrawSlideToImage
            Case "presentation|image": VisualTeX_RedrawPresentationToImage
        End Select
    Else
        Select Case scope
            Case "selection": VisualTeX_RestoreSelectionToLatex
            Case "slide": VisualTeX_RestoreSlideToLatex
            Case "presentation": VisualTeX_RestorePresentationToLatex
        End Select
    End If
    VTPptAgentDocumentOperation = VTPptAgentSessionJson()
End Function

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
    If VTEditSelectedNativeEquation() Then Exit Sub
    VTPowerPointEditShape VTSelectedSingleShape()
    Exit Sub

Failed:
    VTShowError "PowerPoint edit", Err.Number, Err.Description
End Sub

Public Sub VisualTeX_DoubleClickEditSelected()
    ' Invoked by the native macOS double-click monitor for a shape named like a
    ' VisualTeX formula. An error leaving this macro makes PowerPoint show its
    ' run-time error box, so a shape that is not a VisualTeX formula keeps
    ' PowerPoint's own double-click and other failures read as the ribbon's.
    Dim selectedShape As Shape

    On Error GoTo Failed
    VTRequireWritablePowerPointPresentation
    Set selectedShape = VTSelectedSingleShape()
    If Not VTIsVisualTeXPowerPointShape(selectedShape) Then Exit Sub
    VTPowerPointEditShape selectedShape
    Exit Sub

Failed:
    VTShowError "PowerPoint edit", Err.Number, Err.Description
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
    Dim taggedReference As String

    On Error GoTo InvalidShape
    If selectedShape Is Nothing Then Exit Function
    formulaId = VTShapeFormulaId(selectedShape)
    encodedMetadata = VTShapeMetadata(selectedShape)
    If Not VTIsCanonicalUuid(formulaId) Or Not VTIsEncodedMetadata(encodedMetadata) Then Exit Function
    If Not VTTryParseFormulaReference(selectedShape.Title, parsedFormulaId, displayMode, numbered) Then
        ' WPS Presentation cannot set a shape's Title: the formulas it makes
        ' keep the reference in a tag, and their id (name or tag) and metadata
        ' identify them where WPS dropped tags too. VisualTeX checks that the
        ' id and the metadata agree when the editor opens.
        On Error Resume Next
        taggedReference = selectedShape.Tags("VisualTeXFormulaReference")
        Err.Clear
        On Error GoTo InvalidShape
        If Not VTTryParseFormulaReference(taggedReference, parsedFormulaId, displayMode, numbered) Then
            parsedFormulaId = formulaId
        End If
    End If
    If parsedFormulaId <> formulaId Then Exit Function
    VTIsVisualTeXPowerPointShape = True
    Exit Function
InvalidShape:
    VTIsVisualTeXPowerPointShape = False
End Function

Public Function VTPowerPointShapeObjectIdentity( _
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

Public Function VTFindPowerPointShapeById( _
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

    If VTCollectionHasKey(dispatch, "nativeEquation") Then
        If CStr(dispatch("nativeEquation")) = "1" Then
            VTDispatchPowerPointNativeEquation sessionId, dispatch
            Exit Sub
        End If
    End If

    Select Case actionName
        Case "commit": VTFinalizePowerPointDispatch sessionId, dispatch
        Case "cancel": VTCancelPowerPointDispatch sessionId, dispatch
        Case Else
            Err.Raise vbObjectError + 7503, "VisualTeX", "The VisualTeX PowerPoint dispatch action is invalid."
    End Select
    Exit Sub

Failed:
    Dim failureNumber As Long
    Dim failureText As String
    failureNumber = Err.Number
    failureText = Err.Description
    If VTReportCallbackFailure(sessionId, failureNumber, failureText) Then Exit Sub
    Err.Raise failureNumber, "VisualTeX PowerPoint callback", failureText
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

Public Function VTShapeFormulaId(ByVal target As Shape) As String
    Dim value As String
    On Error Resume Next
    value = target.Tags("VisualTeXFormulaId")
    On Error GoTo 0
    If Not VTIsCanonicalUuid(value) And Left$(target.Name, Len(VT_SHAPE_PREFIX)) = VT_SHAPE_PREFIX Then
        value = Mid$(target.Name, Len(VT_SHAPE_PREFIX) + 1)
    End If
    If VTIsCanonicalUuid(value) Then VTShapeFormulaId = value
End Function

Public Function VTShapeMetadata(ByVal target As Shape) As String
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

Public Sub VTSetPowerPointFormulaScaleState( _
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

Public Function VTPowerPointGeometryJson( _
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

Public Function VTPresentationIdentityFor( _
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

Public Function VTFindPowerPointPresentation( _
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

Public Sub VTRequireWritablePowerPointPresentationObject( _
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

Public Sub VTRequireWritablePowerPointPresentation()
    If Presentations.Count = 0 Then
        Err.Raise vbObjectError + 7510, "VisualTeX", "Open a PowerPoint presentation first."
    End If
    VTRequireWritablePowerPointPresentationObject ActivePresentation
    If ActiveWindow Is Nothing Then
        Err.Raise vbObjectError + 7512, "VisualTeX", "Switch PowerPoint to a normal editing view."
    End If
End Sub

Public Sub VTSetShapeTag(ByVal target As Shape, ByVal key As String, ByVal value As String)
    On Error Resume Next
    target.Tags.Delete key
    Err.Clear
    On Error GoTo Failed
    target.Tags.Add key, value
    Exit Sub
Failed:
    Err.Raise vbObjectError + 7513, "VisualTeX", "PowerPoint could not persist VisualTeX tag " & key & "."
End Sub

Public Sub VTRestoreZOrder(ByVal target As Shape, ByVal expectedPosition As Long)
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
