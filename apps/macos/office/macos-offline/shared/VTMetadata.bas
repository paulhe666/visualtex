Attribute VB_Name = "VTMetadata"
Option Explicit

Public Function VTUnicodeText(ParamArray codePoints() As Variant) As String
    Dim index As Long
    Dim codePoint As Long

    For index = LBound(codePoints) To UBound(codePoints)
        codePoint = CLng(codePoints(index))
        If codePoint < 0 Or codePoint > 65535 Then
            Err.Raise vbObjectError + 7214, "VisualTeX", _
                "The requested Ribbon Unicode code point is invalid."
        End If
        If codePoint > 32767 Then codePoint = codePoint - 65536
        VTUnicodeText = VTUnicodeText & ChrW(codePoint)
    Next index
End Function

Public Function VTFormulaFontPresetCount() As Long
    VTFormulaFontPresetCount = 20
End Function

Public Function VTFormulaFontPresetSize( _
    ByVal presetIndex As Long) As Double

    Select Case presetIndex
        Case 0: VTFormulaFontPresetSize = 5#
        Case 1: VTFormulaFontPresetSize = 5.5
        Case 2: VTFormulaFontPresetSize = 6.5
        Case 3: VTFormulaFontPresetSize = 7.5
        Case 4: VTFormulaFontPresetSize = 9#
        Case 5: VTFormulaFontPresetSize = 10.5
        Case 6: VTFormulaFontPresetSize = 12#
        Case 7: VTFormulaFontPresetSize = 14#
        Case 8: VTFormulaFontPresetSize = 15#
        Case 9: VTFormulaFontPresetSize = 16#
        Case 10: VTFormulaFontPresetSize = 18#
        Case 11: VTFormulaFontPresetSize = 22#
        Case 12: VTFormulaFontPresetSize = 24#
        Case 13: VTFormulaFontPresetSize = 26#
        Case 14: VTFormulaFontPresetSize = 36#
        Case 15: VTFormulaFontPresetSize = 42#
        Case 16: VTFormulaFontPresetSize = 48#
        Case 17: VTFormulaFontPresetSize = 54#
        Case 18: VTFormulaFontPresetSize = 72#
        Case 19: VTFormulaFontPresetSize = 96#
        Case Else
            Err.Raise vbObjectError + 7212, "VisualTeX", _
                "The requested formula font-size preset does not exist."
    End Select
End Function

Public Function VTFormulaFontPresetLabel( _
    ByVal presetIndex As Long) As String

    Select Case presetIndex
        Case 0
            VTFormulaFontPresetLabel = _
                VTUnicodeText(20843, 21495) & " (5 pt)"
        Case 1
            VTFormulaFontPresetLabel = _
                VTUnicodeText(19971, 21495) & " (5.5 pt)"
        Case 2
            VTFormulaFontPresetLabel = _
                VTUnicodeText(23567, 20845) & " (6.5 pt)"
        Case 3
            VTFormulaFontPresetLabel = _
                VTUnicodeText(20845, 21495) & " (7.5 pt)"
        Case 4
            VTFormulaFontPresetLabel = _
                VTUnicodeText(23567, 20116) & " (9 pt)"
        Case 5
            VTFormulaFontPresetLabel = _
                VTUnicodeText(20116, 21495) & " (10.5 pt)"
        Case 6
            VTFormulaFontPresetLabel = _
                VTUnicodeText(23567, 22235) & " (12 pt)"
        Case 7
            VTFormulaFontPresetLabel = _
                VTUnicodeText(22235, 21495) & " (14 pt)"
        Case 8
            VTFormulaFontPresetLabel = _
                VTUnicodeText(23567, 19977) & " (15 pt)"
        Case 9
            VTFormulaFontPresetLabel = _
                VTUnicodeText(19977, 21495) & " (16 pt)"
        Case 10
            VTFormulaFontPresetLabel = _
                VTUnicodeText(23567, 20108) & " (18 pt)"
        Case 11
            VTFormulaFontPresetLabel = _
                VTUnicodeText(20108, 21495) & " (22 pt)"
        Case 12
            VTFormulaFontPresetLabel = _
                VTUnicodeText(23567, 19968) & " (24 pt)"
        Case 13
            VTFormulaFontPresetLabel = _
                VTUnicodeText(19968, 21495) & " (26 pt)"
        Case 14
            VTFormulaFontPresetLabel = _
                VTUnicodeText(23567, 21021) & " (36 pt)"
        Case 15
            VTFormulaFontPresetLabel = _
                VTUnicodeText(21021, 21495) & " (42 pt)"
        Case 16
            VTFormulaFontPresetLabel = "48 pt"
        Case 17
            VTFormulaFontPresetLabel = "54 pt"
        Case 18
            VTFormulaFontPresetLabel = "72 pt"
        Case 19
            VTFormulaFontPresetLabel = "96 pt"
        Case Else
            Err.Raise vbObjectError + 7213, "VisualTeX", _
                "The requested formula font-size label does not exist."
    End Select
End Function

Public Function VTFormulaFontPresetIndex( _
    ByVal fontSizePt As Double) As Long

    Dim presetIndex As Long

    VTFormulaFontPresetIndex = -1
    For presetIndex = 0 To VTFormulaFontPresetCount() - 1
        If Abs(fontSizePt - VTFormulaFontPresetSize(presetIndex)) <= 0.05 Then
            VTFormulaFontPresetIndex = presetIndex
            Exit Function
        End If
    Next presetIndex
End Function

Public Function VTFormulaFontSizeDisplayLabel( _
    ByVal fontSizePt As Double) As String

    Dim presetIndex As Long

    presetIndex = VTFormulaFontPresetIndex(fontSizePt)
    If presetIndex >= 0 Then
        VTFormulaFontSizeDisplayLabel = _
            VTFormulaFontPresetLabel(presetIndex)
    ElseIf fontSizePt > 0# Then
        VTFormulaFontSizeDisplayLabel = _
            VTUnicodeText(33258, 23450, 20041) & " (" & _
            VTJsonNumber(fontSizePt) & " pt)"
    Else
        VTFormulaFontSizeDisplayLabel = _
            VTUnicodeText(26410, 36873, 25321, 20844, 24335)
    End If
End Function

Public Function VTIsEncodedMetadata(ByVal value As String) As Boolean
    Dim index As Long
    Dim current As String

    If Len(value) <= Len(VT_METADATA_PREFIX) Or Len(value) > 131072 Then Exit Function
    If Left$(value, Len(VT_METADATA_PREFIX)) <> VT_METADATA_PREFIX Then Exit Function
    For index = Len(VT_METADATA_PREFIX) + 1 To Len(value)
        current = Mid$(value, index, 1)
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_", current, vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next index
    VTIsEncodedMetadata = True
End Function

Public Function VTFormulaReference(ByVal formulaId As String, ByVal displayMode As String, ByVal numbered As Boolean) As String
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7200, "VisualTeX", "Invalid VisualTeX formula id."
    End If
    If displayMode <> "inline" And displayMode <> "block" Then
        Err.Raise vbObjectError + 7201, "VisualTeX", "Invalid VisualTeX display mode."
    End If
    VTFormulaReference = VT_FORMULA_REF_PREFIX & formulaId & ":" & displayMode & ":" & IIf(numbered, "1", "0")
End Function

Public Function VTTryParseFormulaReference(ByVal value As String, ByRef formulaId As String, ByRef displayMode As String, ByRef numbered As Boolean) As Boolean
    Dim payload As String
    Dim fields() As String

    If Left$(value, Len(VT_FORMULA_REF_PREFIX)) <> VT_FORMULA_REF_PREFIX Then Exit Function
    payload = Mid$(value, Len(VT_FORMULA_REF_PREFIX) + 1)
    fields = Split(payload, ":")
    If UBound(fields) <> 2 Then Exit Function
    If Not VTIsCanonicalUuid(fields(0)) Then Exit Function
    If fields(1) <> "inline" And fields(1) <> "block" Then Exit Function
    If fields(2) <> "0" And fields(2) <> "1" Then Exit Function

    formulaId = fields(0)
    displayMode = fields(1)
    numbered = (fields(2) = "1")
    VTTryParseFormulaReference = True
End Function

Public Function VTTryParsePendingMarker(ByVal value As String, ByRef sessionId As String, ByRef formulaId As String) As Boolean
    Dim payload As String
    Dim separator As Long

    If Left$(value, Len(VT_PENDING_PREFIX)) <> VT_PENDING_PREFIX Then Exit Function
    payload = Mid$(value, Len(VT_PENDING_PREFIX) + 1)
    separator = InStr(1, payload, ":", vbBinaryCompare)
    If separator <= 1 Then Exit Function
    sessionId = Left$(payload, separator - 1)
    formulaId = Mid$(payload, separator + 1)
    VTTryParsePendingMarker = VTIsCanonicalUuid(sessionId) And VTIsCanonicalUuid(formulaId)
End Function

Public Sub VTValidateEditEnvelope(ByVal encodedMetadata As String, ByVal formulaReference As String, ByRef formulaId As String, ByRef displayMode As String, ByRef numbered As Boolean)
    If Not VTIsEncodedMetadata(encodedMetadata) Then
        Err.Raise vbObjectError + 7202, "VisualTeX", "The selected object does not contain valid VisualTeX metadata."
    End If

    formulaId = ""
    displayMode = ""
    numbered = False
    If Len(formulaReference) > 0 Then
        If Not VTTryParseFormulaReference(formulaReference, formulaId, displayMode, numbered) Then
            ' Office.js compatibility formulas historically stored the compressed
            ' marker in both Title and AlternativeText. VisualTeX performs the
            ' full inflate/schema/formulaId/lines validation before opening the
            ' editor, so an absent compact reference remains a supported input.
            formulaId = ""
            displayMode = ""
            numbered = False
        End If
    End If
End Sub

Public Function VTRequestJson( _
    ByVal sessionId As String, _
    ByVal hostName As String, _
    ByVal mode As String, _
    ByVal formulaId As String, _
    ByVal displayMode As String, _
    ByVal numbered As Boolean, _
    ByVal sourceDocumentId As String, _
    ByVal sourceObjectId As String, _
    ByVal encodedMetadata As String, _
    ByVal pendingMarker As String, _
    Optional ByVal powerPointJson As String = "", _
    Optional ByVal nativeEquation As Boolean = False, _
    Optional ByVal fontSizePt As Double = 0#, _
    Optional ByVal referenceWidthPt As Double = 0#, _
    Optional ByVal referenceHeightPt As Double = 0#, _
    Optional ByVal operationName As String = "formula", _
    Optional ByVal forkCopiedFormula As Boolean = False) As String

    If Not VTIsCanonicalUuid(sessionId) Then
        Err.Raise vbObjectError + 7203, "VisualTeX", "Invalid VisualTeX Session id."
    End If
    If hostName <> "word" And hostName <> "powerpoint" Then
        Err.Raise vbObjectError + 7204, "VisualTeX", "Invalid VisualTeX Office host."
    End If
    If mode <> "create" And mode <> "edit" Then
        Err.Raise vbObjectError + 7205, "VisualTeX", "Invalid VisualTeX Session mode."
    End If
    If Len(formulaId) > 0 And Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7206, "VisualTeX", "Invalid VisualTeX formula id."
    End If
    If displayMode <> "inline" And displayMode <> "block" Then
        Err.Raise vbObjectError + 7207, "VisualTeX", "Invalid VisualTeX display mode."
    End If
    If numbered And (hostName <> "word" Or displayMode <> "block") Then
        Err.Raise vbObjectError + 7208, "VisualTeX", "Only Word display formulas can be numbered."
    End If
    If Len(encodedMetadata) > 0 And Not VTIsEncodedMetadata(encodedMetadata) Then
        Err.Raise vbObjectError + 7209, "VisualTeX", "Invalid VisualTeX metadata envelope."
    End If
    If fontSizePt < 0# Or referenceWidthPt < 0# Or referenceHeightPt < 0# Then
        Err.Raise vbObjectError + 7210, "VisualTeX", _
            "Word image font-size metadata cannot be negative."
    End If
    If hostName <> "word" And _
       (nativeEquation Or fontSizePt > 0# Or referenceWidthPt > 0# Or _
        referenceHeightPt > 0#) Then
        Err.Raise vbObjectError + 7211, "VisualTeX", _
            "PowerPoint requests cannot contain Word-only formula metadata."
    End If
    If operationName <> "formula" And _
       operationName <> "nativeToImage" And _
       operationName <> "imageToNative" Then
        Err.Raise vbObjectError + 7212, "VisualTeX", _
            "Invalid VisualTeX formula operation."
    End If
    If operationName = "nativeToImage" And _
       (hostName <> "word" Or mode <> "edit" Or nativeEquation) Then
        Err.Raise vbObjectError + 7212, "VisualTeX", _
            "Native-to-image conversion requires a Word image-output edit request."
    End If
    If operationName = "imageToNative" And _
       (hostName <> "word" Or mode <> "edit" Or Not nativeEquation) Then
        Err.Raise vbObjectError + 7212, "VisualTeX", _
            "Image-to-native conversion requires a Word native-output edit request."
    End If

    VTRequestJson = "{" & _
        """protocolVersion"":" & CStr(VT_PROTOCOL_VERSION) & "," & _
        """sessionId"":" & VTJsonString(sessionId) & "," & _
        """host"":" & VTJsonString(hostName) & "," & _
        """mode"":" & VTJsonString(mode) & "," & _
        """operation"":" & VTJsonString(operationName) & "," & _
        """forkCopiedFormula"":" & VTJsonBoolean(forkCopiedFormula) & "," & _
        """formulaId"":" & VTJsonNullableString(formulaId) & "," & _
        """displayMode"":" & VTJsonString(displayMode) & "," & _
        """numbered"":" & VTJsonBoolean(numbered) & "," & _
        """nativeEquation"":" & VTJsonBoolean(nativeEquation) & "," & _
        """sourceDocumentId"":" & VTJsonNullableString(sourceDocumentId) & "," & _
        """sourceObjectId"":" & VTJsonNullableString(sourceObjectId) & "," & _
        """encodedMetadata"":" & VTJsonNullableString(encodedMetadata) & "," & _
        """pendingMarker"":" & VTJsonNullableString(pendingMarker) & "," & _
        """fontSizePt"":" & IIf(fontSizePt > 0#, VTJsonNumber(fontSizePt), "null") & "," & _
        """referenceWidthPt"":" & IIf(referenceWidthPt > 0#, VTJsonNumber(referenceWidthPt), "null") & "," & _
        """referenceHeightPt"":" & IIf(referenceHeightPt > 0#, VTJsonNumber(referenceHeightPt), "null") & "," & _
        """powerPoint"":" & IIf(Len(powerPointJson) = 0, "null", powerPointJson) & _
        "}"
End Function
