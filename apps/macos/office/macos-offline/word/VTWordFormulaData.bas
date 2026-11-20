Attribute VB_Name = "VTWordFormulaData"
Option Explicit
Option Private Module

' The formula data kept in document variables: LaTeX, OMML, metadata and format, chunked and Base64URL
' encoded.

Private Const VT_WORD_LATEX_VARIABLE_PREFIX As String = "VT_Latex_"
Private Const VT_WORD_OMML_VARIABLE_PREFIX As String = "VT_OMML_"
Private Const VT_WORD_METADATA_VARIABLE_PREFIX As String = "VT_Metadata_"
Private Const VT_WORD_FORMAT_VARIABLE_PREFIX As String = "VT_Format_"
Public Const VT_WORD_PAYLOAD_CHUNK_SIZE As Long = 20000
Public Const VT_WORD_PAYLOAD_MAX_CHUNKS As Long = 128
Private Const VT_WORD_LATEX_CHUNK_SIZE As Long = VT_WORD_PAYLOAD_CHUNK_SIZE
Private Const VT_WORD_LATEX_MAX_CHUNKS As Long = VT_WORD_PAYLOAD_MAX_CHUNKS

Public Function VTTryFormulaIdForMetadataPayload( _
    ByVal documentObject As Document, _
    ByVal encodedMetadata As String, _
    ByRef formulaId As String) As Boolean

    Dim candidateVariable As Object
    Dim variableName As String
    Dim compactFormulaId As String
    Dim candidateFormulaId As String
    Dim storedMetadata As String
    Dim matchedFormulaId As String
    Dim metadataReadable As Boolean
    Dim matchCount As Long
    Dim prefixLength As Long
    Dim suffixLength As Long

    On Error GoTo ResolveFailed
    formulaId = ""
    If documentObject Is Nothing Or _
       Not VTIsEncodedMetadata(encodedMetadata) Then Exit Function
    prefixLength = Len(VT_WORD_METADATA_VARIABLE_PREFIX)
    suffixLength = Len("_Count")

    For Each candidateVariable In documentObject.Variables
        variableName = CStr(candidateVariable.Name)
        If Len(variableName) > prefixLength + suffixLength And _
           Left$(variableName, prefixLength) = _
               VT_WORD_METADATA_VARIABLE_PREFIX And _
           Right$(variableName, suffixLength) = "_Count" Then
            compactFormulaId = Mid$( _
                variableName, prefixLength + 1, _
                Len(variableName) - prefixLength - suffixLength)
            candidateFormulaId = Replace$(compactFormulaId, "_", "-")
            If VTIsCanonicalUuid(candidateFormulaId) Then
                storedMetadata = ""
                metadataReadable = False
                On Error Resume Next
                Err.Clear
                metadataReadable = VTTryReadWordMetadataPayload( _
                    documentObject, candidateFormulaId, storedMetadata)
                Err.Clear
                On Error GoTo ResolveFailed
                If metadataReadable And _
                   StrComp(storedMetadata, encodedMetadata, _
                       vbBinaryCompare) = 0 Then
                    matchCount = matchCount + 1
                    matchedFormulaId = candidateFormulaId
                End If
            End If
        End If
    Next candidateVariable

    If matchCount = 1 Then
        formulaId = matchedFormulaId
        VTTryFormulaIdForMetadataPayload = True
    End If
    Exit Function

ResolveFailed:
    formulaId = ""
    VTTryFormulaIdForMetadataPayload = False
End Function

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

Public Function VTIsBase64UrlPayload(ByVal value As String) As Boolean
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

Private Function VTStoredPayloadChunkCount( _
    ByVal documentObject As Document, _
    ByVal countVariableName As String, _
    ByVal firstChunkVariableName As String, _
    ByVal maximumChunks As Long) As Long

    Dim countText As String
    Dim ignoredChunk As String
    Dim parsedCount As Long

    If documentObject Is Nothing Or maximumChunks < 1 Then Exit Function
    If VTTryGetDocumentVariable( _
       documentObject, countVariableName, countText) Then
        If Len(countText) > 0 And IsNumeric(countText) Then
            On Error GoTo InvalidCount
            parsedCount = CLng(countText)
            On Error GoTo 0
            If parsedCount >= 1 And parsedCount <= maximumChunks Then
                VTStoredPayloadChunkCount = parsedCount
                Exit Function
            End If
        End If
        ' A corrupt published count requires a complete cleanup, but this slow
        ' path is never used by a healthy formula edit.
        VTStoredPayloadChunkCount = maximumChunks
        Exit Function
    End If
    If VTTryGetDocumentVariable( _
       documentObject, firstChunkVariableName, ignoredChunk) Then
        ' Recover orphan chunks left by an interrupted legacy write.
        VTStoredPayloadChunkCount = maximumChunks
    End If
    Exit Function

InvalidCount:
    Err.Clear
    On Error GoTo 0
    VTStoredPayloadChunkCount = maximumChunks
End Function

Public Sub VTDeleteWordLatexPayload(ByVal documentObject As Document, ByVal formulaId As String)
    Dim index As Long
    Dim chunkCount As Long

    chunkCount = VTStoredPayloadChunkCount( _
        documentObject, _
        VTWordLatexCountVariableName(formulaId), _
        VTWordLatexChunkVariableName(formulaId, 1), _
        VT_WORD_LATEX_MAX_CHUNKS)
    For index = 1 To chunkCount
        VTDeleteDocumentVariable documentObject, VTWordLatexChunkVariableName(formulaId, index)
    Next index
    VTDeleteDocumentVariable documentObject, VTWordLatexCountVariableName(formulaId)
End Sub

Public Sub VTSetWordLatexPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal latexBase64 As String)

    Dim chunkCount As Long
    Dim previousChunkCount As Long
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

    previousChunkCount = VTStoredPayloadChunkCount( _
        documentObject, _
        VTWordLatexCountVariableName(formulaId), _
        VTWordLatexChunkVariableName(formulaId, 1), _
        VT_WORD_LATEX_MAX_CHUNKS)
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
    For index = chunkCount + 1 To previousChunkCount
        VTDeleteDocumentVariable _
            documentObject, VTWordLatexChunkVariableName(formulaId, index)
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

Public Function VTTryReadWordLatexPayload( _
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

Private Function VTWordOmmlVariableStem(ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7471, "VisualTeX", "VisualTeX cannot address Word OMML metadata for an invalid formula id."
    End If
    VTWordOmmlVariableStem = _
        VT_WORD_OMML_VARIABLE_PREFIX & Replace$(formulaId, "-", "_")
End Function

Private Function VTWordOmmlCountVariableName(ByVal formulaId As String) As String
    VTWordOmmlCountVariableName = VTWordOmmlVariableStem(formulaId) & "_Count"
End Function

Private Function VTWordOmmlChunkVariableName( _
    ByVal formulaId As String, _
    ByVal index As Long) As String

    If index < 1 Or index > VT_WORD_OMML_MAX_CHUNKS Then
        Err.Raise vbObjectError + 7472, "VisualTeX", "VisualTeX Word OMML metadata chunk index is invalid."
    End If
    VTWordOmmlChunkVariableName = _
        VTWordOmmlVariableStem(formulaId) & "_" & Right$("000" & CStr(index), 3)
End Function

Public Sub VTDeleteWordOmmlPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String)

    Dim index As Long
    Dim chunkCount As Long

    chunkCount = VTStoredPayloadChunkCount( _
        documentObject, _
        VTWordOmmlCountVariableName(formulaId), _
        VTWordOmmlChunkVariableName(formulaId, 1), _
        VT_WORD_OMML_MAX_CHUNKS)
    For index = 1 To chunkCount
        VTDeleteDocumentVariable _
            documentObject, VTWordOmmlChunkVariableName(formulaId, index)
    Next index
    VTDeleteDocumentVariable documentObject, VTWordOmmlCountVariableName(formulaId)
End Sub

Public Sub VTSetWordOmmlPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal ommlBase64 As String)

    Dim chunkCount As Long
    Dim previousChunkCount As Long
    Dim index As Long
    Dim chunkValue As String
    Dim storageErrorNumber As Long
    Dim storageErrorDescription As String

    If Not VTIsBase64UrlPayload(ommlBase64) Then
        Err.Raise vbObjectError + 7473, "VisualTeX", "VisualTeX Word OMML metadata is invalid or too large."
    End If
    chunkCount = _
        (Len(ommlBase64) + VT_WORD_OMML_CHUNK_SIZE - 1) \ _
        VT_WORD_OMML_CHUNK_SIZE
    If chunkCount < 1 Or chunkCount > VT_WORD_OMML_MAX_CHUNKS Then
        Err.Raise vbObjectError + 7473, "VisualTeX", "VisualTeX Word OMML metadata requires too many chunks."
    End If

    previousChunkCount = VTStoredPayloadChunkCount( _
        documentObject, _
        VTWordOmmlCountVariableName(formulaId), _
        VTWordOmmlChunkVariableName(formulaId, 1), _
        VT_WORD_OMML_MAX_CHUNKS)
    On Error GoTo StorageFailed
    For index = 1 To chunkCount
        chunkValue = Mid$( _
            ommlBase64, _
            (index - 1) * VT_WORD_OMML_CHUNK_SIZE + 1, _
            VT_WORD_OMML_CHUNK_SIZE)
        VTSetDocumentVariable _
            documentObject, _
            VTWordOmmlChunkVariableName(formulaId, index), _
            chunkValue
    Next index
    For index = chunkCount + 1 To previousChunkCount
        VTDeleteDocumentVariable _
            documentObject, VTWordOmmlChunkVariableName(formulaId, index)
    Next index
    VTSetDocumentVariable _
        documentObject, _
        VTWordOmmlCountVariableName(formulaId), _
        CStr(chunkCount)
    Exit Sub

StorageFailed:
    storageErrorNumber = Err.Number
    storageErrorDescription = Err.Description
    On Error Resume Next
    VTDeleteWordOmmlPayload documentObject, formulaId
    On Error GoTo 0
    Err.Raise storageErrorNumber, "VisualTeX Word OMML metadata", storageErrorDescription
End Sub

Public Function VTTryReadWordOmmlPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByRef ommlBase64 As String) As Boolean

    Dim countText As String
    Dim chunkValue As String
    Dim chunkCount As Long
    Dim index As Long
    Dim ommlXml As String

    ommlBase64 = ""
    If Not VTTryGetDocumentVariable( _
        documentObject, VTWordOmmlCountVariableName(formulaId), countText) Then Exit Function
    If Len(countText) = 0 Or Not IsNumeric(countText) Then GoTo InvalidPayload
    chunkCount = CLng(countText)
    If chunkCount < 1 Or chunkCount > VT_WORD_OMML_MAX_CHUNKS Then GoTo InvalidPayload

    For index = 1 To chunkCount
        If Not VTTryGetDocumentVariable( _
            documentObject, _
            VTWordOmmlChunkVariableName(formulaId, index), _
            chunkValue) Then GoTo InvalidPayload
        ommlBase64 = ommlBase64 & chunkValue
    Next index
    If Not VTIsBase64UrlPayload(ommlBase64) Then GoTo InvalidPayload
    ommlXml = VTBase64UrlDecodeUtf8(ommlBase64)
    VTValidateOmmlFragment ommlXml
    VTTryReadWordOmmlPayload = True
    Exit Function

InvalidPayload:
    Err.Raise vbObjectError + 7474, "VisualTeX", "The stored Word OMML metadata is incomplete or corrupt."
End Function

Private Function VTWordMetadataVariableStem(ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7463, "VisualTeX", "VisualTeX cannot address Word formula metadata for an invalid formula id."
    End If
    VTWordMetadataVariableStem = _
        VT_WORD_METADATA_VARIABLE_PREFIX & Replace$(formulaId, "-", "_")
End Function

Private Function VTWordMetadataCountVariableName(ByVal formulaId As String) As String
    VTWordMetadataCountVariableName = VTWordMetadataVariableStem(formulaId) & "_Count"
End Function

Private Function VTWordMetadataChunkVariableName( _
    ByVal formulaId As String, _
    ByVal index As Long) As String

    If index < 1 Or index > VT_WORD_PAYLOAD_MAX_CHUNKS Then
        Err.Raise vbObjectError + 7464, "VisualTeX", "VisualTeX Word metadata chunk index is invalid."
    End If
    VTWordMetadataChunkVariableName = _
        VTWordMetadataVariableStem(formulaId) & "_" & Right$("000" & CStr(index), 3)
End Function

Public Sub VTDeleteWordMetadataPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String)

    Dim index As Long
    Dim chunkCount As Long

    chunkCount = VTStoredPayloadChunkCount( _
        documentObject, _
        VTWordMetadataCountVariableName(formulaId), _
        VTWordMetadataChunkVariableName(formulaId, 1), _
        VT_WORD_PAYLOAD_MAX_CHUNKS)
    For index = 1 To chunkCount
        VTDeleteDocumentVariable _
            documentObject, VTWordMetadataChunkVariableName(formulaId, index)
    Next index
    VTDeleteDocumentVariable documentObject, VTWordMetadataCountVariableName(formulaId)
End Sub

Public Sub VTSetWordMetadataPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal encodedMetadata As String)

    Dim chunkCount As Long
    Dim previousChunkCount As Long
    Dim index As Long
    Dim chunkValue As String
    Dim storageErrorNumber As Long
    Dim storageErrorDescription As String

    If Not VTIsEncodedMetadata(encodedMetadata) Then
        Err.Raise vbObjectError + 7465, "VisualTeX", "VisualTeX Word formula metadata is invalid or too large."
    End If
    chunkCount = _
        (Len(encodedMetadata) + VT_WORD_PAYLOAD_CHUNK_SIZE - 1) \ _
        VT_WORD_PAYLOAD_CHUNK_SIZE
    If chunkCount < 1 Or chunkCount > VT_WORD_PAYLOAD_MAX_CHUNKS Then
        Err.Raise vbObjectError + 7465, "VisualTeX", "VisualTeX Word formula metadata requires too many chunks."
    End If

    previousChunkCount = VTStoredPayloadChunkCount( _
        documentObject, _
        VTWordMetadataCountVariableName(formulaId), _
        VTWordMetadataChunkVariableName(formulaId, 1), _
        VT_WORD_PAYLOAD_MAX_CHUNKS)
    On Error GoTo StorageFailed
    For index = 1 To chunkCount
        chunkValue = Mid$( _
            encodedMetadata, _
            (index - 1) * VT_WORD_PAYLOAD_CHUNK_SIZE + 1, _
            VT_WORD_PAYLOAD_CHUNK_SIZE)
        VTSetDocumentVariable _
            documentObject, _
            VTWordMetadataChunkVariableName(formulaId, index), _
            chunkValue
    Next index
    For index = chunkCount + 1 To previousChunkCount
        VTDeleteDocumentVariable _
            documentObject, VTWordMetadataChunkVariableName(formulaId, index)
    Next index
    VTSetDocumentVariable _
        documentObject, _
        VTWordMetadataCountVariableName(formulaId), _
        CStr(chunkCount)
    Exit Sub

StorageFailed:
    storageErrorNumber = Err.Number
    storageErrorDescription = Err.Description
    On Error Resume Next
    VTDeleteWordMetadataPayload documentObject, formulaId
    On Error GoTo 0
    Err.Raise storageErrorNumber, "VisualTeX Word formula metadata", storageErrorDescription
End Sub

Public Function VTTryReadWordMetadataPayload( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByRef encodedMetadata As String) As Boolean

    Dim countText As String
    Dim chunkValue As String
    Dim chunkCount As Long
    Dim index As Long

    encodedMetadata = ""
    If Not VTTryGetDocumentVariable( _
        documentObject, VTWordMetadataCountVariableName(formulaId), countText) Then Exit Function
    If Len(countText) = 0 Or Not IsNumeric(countText) Then GoTo InvalidPayload
    chunkCount = CLng(countText)
    If chunkCount < 1 Or chunkCount > VT_WORD_PAYLOAD_MAX_CHUNKS Then GoTo InvalidPayload

    For index = 1 To chunkCount
        If Not VTTryGetDocumentVariable( _
            documentObject, _
            VTWordMetadataChunkVariableName(formulaId, index), _
            chunkValue) Then GoTo InvalidPayload
        encodedMetadata = encodedMetadata & chunkValue
    Next index
    If Not VTIsEncodedMetadata(encodedMetadata) Then GoTo InvalidPayload
    VTTryReadWordMetadataPayload = True
    Exit Function

InvalidPayload:
    Err.Raise vbObjectError + 7466, "VisualTeX", "The stored Word formula metadata is incomplete or corrupt."
End Function

Public Function VTWordFormatVariableName(ByVal formulaId As String) As String
    If Not VTIsCanonicalUuid(formulaId) Then
        Err.Raise vbObjectError + 7467, "VisualTeX", "VisualTeX cannot address Word formula format for an invalid formula id."
    End If
    VTWordFormatVariableName = _
        VT_WORD_FORMAT_VARIABLE_PREFIX & Replace$(formulaId, "-", "_")
End Function

Public Sub VTSetWordFormulaFormat( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByVal displayMode As String, _
    ByVal numbered As Boolean)

    If displayMode <> "inline" And displayMode <> "block" Then
        Err.Raise vbObjectError + 7468, "VisualTeX", "VisualTeX Word formula format has an invalid display mode."
    End If
    If numbered And displayMode <> "block" Then
        Err.Raise vbObjectError + 7468, "VisualTeX", "Only display formulas can retain Word equation numbers."
    End If
    VTSetDocumentVariable _
        documentObject, _
        VTWordFormatVariableName(formulaId), _
        displayMode & "|" & IIf(numbered, "1", "0")
End Sub

Public Function VTTryReadWordFormulaFormat( _
    ByVal documentObject As Document, _
    ByVal formulaId As String, _
    ByRef displayMode As String, _
    ByRef numbered As Boolean) As Boolean

    Dim storedValue As String
    Dim fields() As String

    displayMode = ""
    numbered = False
    If Not VTTryGetDocumentVariable( _
        documentObject, VTWordFormatVariableName(formulaId), storedValue) Then Exit Function
    fields = Split(storedValue, "|")
    If UBound(fields) <> 1 Then GoTo InvalidFormat
    displayMode = fields(0)
    If displayMode <> "inline" And displayMode <> "block" Then GoTo InvalidFormat
    If fields(1) = "1" Then
        numbered = True
    ElseIf fields(1) <> "0" Then
        GoTo InvalidFormat
    End If
    If numbered And displayMode <> "block" Then GoTo InvalidFormat
    VTTryReadWordFormulaFormat = True
    Exit Function

InvalidFormat:
    Err.Raise vbObjectError + 7469, "VisualTeX", "The stored Word formula format is invalid or corrupt."
End Function
