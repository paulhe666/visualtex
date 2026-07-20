Attribute VB_Name = "VTMetadata"
Option Explicit

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
    Optional ByVal powerPointJson As String = "") As String

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

    VTRequestJson = "{" & _
        """protocolVersion"":" & CStr(VT_PROTOCOL_VERSION) & "," & _
        """sessionId"":" & VTJsonString(sessionId) & "," & _
        """host"":" & VTJsonString(hostName) & "," & _
        """mode"":" & VTJsonString(mode) & "," & _
        """formulaId"":" & VTJsonNullableString(formulaId) & "," & _
        """displayMode"":" & VTJsonString(displayMode) & "," & _
        """numbered"":" & VTJsonBoolean(numbered) & "," & _
        """sourceDocumentId"":" & VTJsonNullableString(sourceDocumentId) & "," & _
        """sourceObjectId"":" & VTJsonNullableString(sourceObjectId) & "," & _
        """encodedMetadata"":" & VTJsonNullableString(encodedMetadata) & "," & _
        """pendingMarker"":" & VTJsonNullableString(pendingMarker) & "," & _
        """powerPoint"":" & IIf(Len(powerPointJson) = 0, "null", powerPointJson) & _
        "}"
End Function
