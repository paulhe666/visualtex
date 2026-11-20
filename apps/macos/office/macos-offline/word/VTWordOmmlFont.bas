Attribute VB_Name = "VTWordOmmlFont"
Option Explicit
Option Private Module

' =====================================================================
' OMML formula font (Windows: the OMML font menu of the font size group)
'
' The chosen font (Cambria Math, STIX Two Math, Latin Modern Math or TeX
' Gyre Termes Math; "" follows the document) is saved for every document.
' Choosing one makes it the active document's math font
' (Document.OMathFontName) and rewrites the explicit fonts of its equations,
' as one Undo step. Every later VisualTeX OMML write first makes it the
' document math font (the app writes in the document math font). A font Word
' does not list is never applied. VisualTeX installs the TrueType STIX Two
' Math, VisualTeX Latin Modern Math and VisualTeX TeX Gyre Termes Math in
' ~/Library/Fonts with its Office add-ins; Word lists them once restarted.
' =====================================================================

Private Const VT_WORD_OMML_FONT_CONTROL_PREFIX As String = _
    "VisualTeX.Mac.Word.OmmlFont."
Private Const VT_WORD_OMML_FONT_FILE As String = "/WordOmmlMathFont.txt"
' The TrueType math fonts VisualTeX installs in ~/Library/Fonts with its
' Office add-ins (office/macos-offline/fonts; the Windows add-in uses the same
' families). A choice saved before 1.3.0 can name the CFF Latin Modern Math.
Private Const VT_OMML_FONT_LATIN_MODERN As String = "VisualTeX Latin Modern Math"
Private Const VT_OMML_FONT_LEGACY_LATIN_MODERN As String = "Latin Modern Math"
Private Const VT_OMML_FONT_TERMES As String = "VisualTeX TeX Gyre Termes Math"
' Equations per staging package when Word equations change their font.
Private Const VT_WORD_MATH_FONT_BATCH_ITEMS As Long = 40
Private Const VT_WORD_MATH_FONT_BATCH_CHARACTERS As Long = 400000
' The OMML formula font (ribbon "OMML font"): "" follows each document's
' math font; otherwise every VisualTeX OMML write and Update Equation Format
' make it the document math font. Saved for every document.
Private VT_WORD_OMML_FONT_LOADED As Boolean
' Lower-case names of the fonts Word lists (Application.FontNames), read once.
Private VT_WORD_FONT_NAMES As Collection
Private VT_WORD_FONT_NAMES_UNAVAILABLE As Boolean

Public Function VTWordMathFontNameFromRange( _
    ByVal sourceRange As Range) As String

    Dim candidate As String

    If sourceRange Is Nothing Then Exit Function
    On Error Resume Next
    candidate = Trim$(CStr(sourceRange.Font.NameAscii))
    If Len(candidate) = 0 Or candidate = "9999999" Then
        candidate = Trim$(CStr(sourceRange.Font.Name))
    End If
    If Len(candidate) = 0 Or candidate = "9999999" Then
        candidate = Trim$(CStr(sourceRange.Font.NameOther))
    End If
    Err.Clear
    On Error GoTo 0

    If Len(candidate) = 0 Or Len(candidate) > 128 Or _
       candidate = "9999999" Or InStr(candidate, vbCr) > 0 Or _
       InStr(candidate, vbLf) > 0 Then Exit Function
    VTWordMathFontNameFromRange = candidate
End Function

Private Sub VTApplyWordNativeMathFontName( _
    ByVal equationRange As Range, _
    ByVal fontName As String)

    Dim trimmedName As String
    Dim errorNumber As Long
    Dim errorDescription As String

    trimmedName = Trim$(fontName)
    If Len(trimmedName) = 0 Then Exit Sub
    If equationRange Is Nothing Or Len(trimmedName) > 128 Or _
       InStr(trimmedName, vbCr) > 0 Or InStr(trimmedName, vbLf) > 0 Then
        Err.Raise vbObjectError + 7534, "VisualTeX", _
            "The preserved Word math font is invalid."
    End If

    On Error GoTo ApplyFailed
    equationRange.Font.Name = trimmedName
    equationRange.Font.NameAscii = trimmedName
    equationRange.Font.NameOther = trimmedName
    Exit Sub

ApplyFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    If errorNumber = 0 Then errorNumber = vbObjectError + 7534
    Err.Raise errorNumber, "VisualTeX Word math font", errorDescription
End Sub

' Word equations are written in the document's math font (Windows passes
' Document.OMathFontName to ApplyExplicitTransferMathFont). Every Word editor
' request carries it; the app builds the equation package with it. A chosen
' OMML font becomes the document math font first.
Public Function VTV2WithDocumentMathFont( _
    ByVal requestJson As String, ByVal documentObject As Document) As String

    Dim fontName As String

    VTV2WithDocumentMathFont = requestJson
    If documentObject Is Nothing Or Right$(requestJson, 1) <> "}" Then Exit Function
    VTApplyChosenOmmlFontToDocument documentObject
    On Error Resume Next
    fontName = Trim$(CStr(documentObject.OMathFontName))
    Err.Clear
    On Error GoTo 0
    If Len(fontName) = 0 Or Len(fontName) > 128 Or _
       InStr(fontName, vbCr) > 0 Or InStr(fontName, vbLf) > 0 Or _
       InStr(fontName, vbTab) > 0 Then Exit Function
    VTV2WithDocumentMathFont = Left$(requestJson, Len(requestJson) - 1) & _
        ",""wordDocumentMathFontName"":" & VTJsonString(fontName) & "}"
End Function

' The menu item (Cambria, Stix, LatinModern, Termes) of an OMML font control;
' "" for Follow.

Public Function VTOmmlControlChoiceKey(ByVal controlId As String) As String
    Select Case controlId
        Case VT_WORD_OMML_FONT_CONTROL_PREFIX & "Cambria"
            VTOmmlControlChoiceKey = "Cambria"
        Case VT_WORD_OMML_FONT_CONTROL_PREFIX & "Stix"
            VTOmmlControlChoiceKey = "Stix"
        Case VT_WORD_OMML_FONT_CONTROL_PREFIX & "LatinModern"
            VTOmmlControlChoiceKey = "LatinModern"
        Case VT_WORD_OMML_FONT_CONTROL_PREFIX & "Termes"
            VTOmmlControlChoiceKey = "Termes"
    End Select
End Function

' The menu item of a saved OMML font: the legacy CFF Latin Modern Math is the
' Latin Modern item. "" when fontName is no menu font.
Public Function VTOmmlFontChoiceKey(ByVal fontName As String) As String
    Select Case fontName
        Case "Cambria Math"
            VTOmmlFontChoiceKey = "Cambria"
        Case "STIX Two Math"
            VTOmmlFontChoiceKey = "Stix"
        Case VT_OMML_FONT_LATIN_MODERN, VT_OMML_FONT_LEGACY_LATIN_MODERN
            VTOmmlFontChoiceKey = "LatinModern"
        Case VT_OMML_FONT_TERMES
            VTOmmlFontChoiceKey = "Termes"
    End Select
End Function

' The font a menu item applies. Latin Modern is the TrueType VisualTeX Latin
' Modern Math (Word's PDF export embeds it); the CFF original only when Word
' lists that one alone.
Public Function VTOmmlFontNameForChoice(ByVal choiceKey As String) As String
    Select Case choiceKey
        Case "Cambria"
            VTOmmlFontNameForChoice = "Cambria Math"
        Case "Stix"
            VTOmmlFontNameForChoice = "STIX Two Math"
        Case "LatinModern"
            VTOmmlFontNameForChoice = VT_OMML_FONT_LATIN_MODERN
            If Not VTWordFontInstalled(VT_OMML_FONT_LATIN_MODERN) Then
                If VTWordFontInstalled(VT_OMML_FONT_LEGACY_LATIN_MODERN) Then _
                    VTOmmlFontNameForChoice = VT_OMML_FONT_LEGACY_LATIN_MODERN
            End If
        Case "Termes"
            VTOmmlFontNameForChoice = VT_OMML_FONT_TERMES
    End Select
End Function

' The menu label of an item (the font's own name, without "VisualTeX").
Public Function VTOmmlFontLabelForChoice(ByVal choiceKey As String) As String
    Select Case choiceKey
        Case "Cambria"
            VTOmmlFontLabelForChoice = "Cambria Math"
        Case "Stix"
            VTOmmlFontLabelForChoice = "STIX Two Math"
        Case "LatinModern"
            VTOmmlFontLabelForChoice = "Latin Modern Math"
        Case "Termes"
            VTOmmlFontLabelForChoice = "TeX Gyre Termes Math"
    End Select
End Function

Public Function VTOmmlFontNameForControl(ByVal controlId As String) As String
    VTOmmlFontNameForControl = _
        VTOmmlFontNameForChoice(VTOmmlControlChoiceKey(controlId))
End Function

Private Function VTIsOmmlFontChoice(ByVal fontName As String) As Boolean
    VTIsOmmlFontChoice = Len(VTOmmlFontChoiceKey(fontName)) > 0
End Function

' The saved choice, read once per Word session. A missing or unreadable file
' follows the document.
Public Sub VTLoadOmmlFontSetting()
    Dim settingPath As String
    Dim value As String

    If VT_WORD_OMML_FONT_LOADED Then Exit Sub
    VT_WORD_OMML_FONT_LOADED = True
    VT_WORD_OMML_FONT_NAME = ""
    On Error GoTo SettingUnavailable
    settingPath = VTApplicationSupportRoot() & VT_WORD_OMML_FONT_FILE
    If Not VTPathFileExists(settingPath) Then Exit Sub
    value = VTReadText(settingPath, 128)
    value = Trim$(Replace$(Replace$(value, vbCr, ""), vbLf, ""))
    If VTIsOmmlFontChoice(value) Then VT_WORD_OMML_FONT_NAME = value
    Exit Sub

SettingUnavailable:
    VT_WORD_OMML_FONT_NAME = ""
    Err.Clear
End Sub

Public Sub VTWriteOmmlFontSetting(ByVal fontName As String)
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo WriteFailed
    If Len(fontName) = 0 Then
        VTWriteTextAtomic VTApplicationSupportRoot() & VT_WORD_OMML_FONT_FILE, "follow"
    Else
        VTWriteTextAtomic VTApplicationSupportRoot() & VT_WORD_OMML_FONT_FILE, fontName
    End If
    VT_WORD_OMML_FONT_LOADED = True
    VT_WORD_OMML_FONT_NAME = fontName
    Exit Sub

WriteFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    Err.Raise errorNumber, "VisualTeX OMML font", _
        "Unable to save the OMML font: " & errorDescription
End Sub

' True when Word lists the font (or cannot list its fonts at all).
Public Function VTWordFontInstalled(ByVal fontName As String) As Boolean
    Dim names As Collection
    Dim candidate As Variant
    Dim key As String

    If VT_WORD_FONT_NAMES Is Nothing And Not VT_WORD_FONT_NAMES_UNAVAILABLE Then
        Set names = New Collection
        On Error GoTo ListUnavailable
        For Each candidate In Application.FontNames
            key = LCase$(Trim$(CStr(candidate)))
            If Len(key) > 0 Then
                If Not VTCollectionHasKey(names, key) Then names.Add True, key
            End If
        Next candidate
        On Error GoTo 0
        If names.Count = 0 Then
            VT_WORD_FONT_NAMES_UNAVAILABLE = True
        Else
            Set VT_WORD_FONT_NAMES = names
        End If
    End If
    If VT_WORD_FONT_NAMES_UNAVAILABLE Then
        VTWordFontInstalled = True
    Else
        VTWordFontInstalled = VTCollectionHasKey( _
            VT_WORD_FONT_NAMES, LCase$(Trim$(fontName)))
    End If
    Exit Function

ListUnavailable:
    Err.Clear
    VT_WORD_FONT_NAMES_UNAVAILABLE = True
    VTWordFontInstalled = True
End Function

' The font of the chosen OMML menu item when Word has it, or "" (follow the
' document). A saved Latin Modern Math choice uses VisualTeX Latin Modern Math
' once Word lists it.
Public Function VTChosenOmmlFontName() As String
    Dim fontName As String

    VTLoadOmmlFontSetting
    If Len(VT_WORD_OMML_FONT_NAME) = 0 Then Exit Function
    fontName = VTOmmlFontNameForChoice(VTOmmlFontChoiceKey(VT_WORD_OMML_FONT_NAME))
    If Len(fontName) = 0 Then Exit Function
    If VTWordFontInstalled(fontName) Then VTChosenOmmlFontName = fontName
End Function

' Document.OMathFontName (Cambria Math when Word reports none).
Public Function VTDocumentMathFontName(ByVal documentObject As Document) As String
    On Error Resume Next
    VTDocumentMathFontName = Trim$(CStr(documentObject.OMathFontName))
    Err.Clear
    On Error GoTo 0
    If Len(VTDocumentMathFontName) = 0 Then VTDocumentMathFontName = "Cambria Math"
End Function

' Before VisualTeX writes OMML into a document, the chosen OMML font becomes
' its math font, so the written equation names it. Never fails the write.
Public Sub VTApplyChosenOmmlFontToDocument(ByVal documentObject As Document)
    Dim fontName As String

    If documentObject Is Nothing Then Exit Sub
    On Error Resume Next
    fontName = VTChosenOmmlFontName()
    If Len(fontName) > 0 Then
        If StrComp(VTDocumentMathFontName(documentObject), fontName, _
           vbTextCompare) <> 0 Then documentObject.OMathFontName = fontName
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub VTInvalidateOmmlFontControls()
    Dim suffix As Variant

    On Error Resume Next
    If VT_WORD_RIBBON Is Nothing Then Exit Sub
    For Each suffix In Array("Follow", "Cambria", "Stix", "LatinModern", "Termes")
        VT_WORD_RIBBON.InvalidateControl VT_WORD_OMML_FONT_CONTROL_PREFIX & CStr(suffix)
    Next suffix
    Err.Clear
    On Error GoTo 0
End Sub

' Makes fontName the document math font and every equation follow it, as one
' Undo step. Returns how many equations changed.
Public Function VTApplyOmmlFontToDocument( _
    ByVal documentObject As Document, _
    ByVal fontName As String) As Long

    Dim undoRecord As Object
    Dim undoRecordStarted As Boolean
    Dim internalMutationStarted As Boolean
    Dim previousScreenUpdating As Boolean
    Dim changed As Long
    Dim errorNumber As Long
    Dim errorDescription As String

    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
    undoRecord.StartCustomRecord "VisualTeX OMML Font"
    undoRecordStarted = True
    VTBeginWordInternalMutation
    internalMutationStarted = True
    If StrComp(VTDocumentMathFontName(documentObject), fontName, _
       vbTextCompare) <> 0 Then documentObject.OMathFontName = fontName
    changed = VTV2MathFontRepairs(documentObject, fontName, True)
    VTEndWordInternalMutation
    internalMutationStarted = False
    undoRecord.EndCustomRecord
    undoRecordStarted = False
    Application.ScreenUpdating = previousScreenUpdating
    VTApplyOmmlFontToDocument = changed
    Exit Function

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If internalMutationStarted Then VTEndWordInternalMutation
    If undoRecordStarted Then
        undoRecord.EndCustomRecord
        documentObject.Undo 1
    End If
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    Err.Raise errorNumber, "VisualTeX OMML font", errorDescription
End Function

' Word equations whose explicit math run fonts are not fontName: every
' w:rFonts of a math run (m:nor normal text, e.g. \mathrm or \text, keeps its
' text font) and of every control property names fontName in w:ascii and
' w:hAnsi (Windows ApplyExplicitTransferMathFont). The number of a numbered
' equation (the equation array, its # and the SEQ fields after it) is left
' as it is. With apply, each such equation is rewritten from a staging
' package and keeps its type, text, fields, bookmarks and metadata. Returns
' how many equations do (or did) not follow fontName.
Public Function VTV2MathFontRepairs( _
    ByVal doc As Document, _
    ByVal fontName As String, _
    ByVal apply As Boolean) As Long

    Dim escapedFont As String
    Dim index As Long
    Dim root As OMath
    Dim exact As Range
    Dim lastStart As Long
    Dim packageXml As String
    Dim bodyXml As String
    Dim rewritten As String
    Dim bodyStart As Long
    Dim bodyEnd As Long
    Dim skeleton As String
    Dim skeletonBodyStart As Long
    Dim skeletonBodyEnd As Long
    Dim starts() As Long
    Dim ends() As Long
    Dim types() As Long
    Dim bodies() As String
    Dim pending As Long
    Dim pendingCharacters As Long
    Dim changed As Long
    Dim numbered As Boolean

    fontName = Trim$(fontName)
    If Len(fontName) = 0 Or doc Is Nothing Then Exit Function
    escapedFont = VTXmlAttributeValue(fontName)
    lastStart = -1
    ' Last to first: a rewritten equation keeps its length, so earlier
    ' equations keep their positions and indices.
    For index = doc.OMaths.Count To 1 Step -1
        Set root = VTNativeMathRoot(doc.OMaths(index))
        Set exact = root.Range.Duplicate
        If lastStart < 0 Or exact.Start < lastStart Then
            lastStart = exact.Start
            packageXml = exact.WordOpenXML
            If VTMathFontPackageBody(packageXml, bodyStart, bodyEnd) Then
                bodyXml = Mid$(packageXml, bodyStart, bodyEnd - bodyStart)
                numbered = False
                If InStr(1, bodyXml, "<m:eqArr", vbBinaryCompare) > 0 Then _
                    numbered = Not VTPureNativeNumberField(exact) Is Nothing
                rewritten = VTMathFontRewriteBody(bodyXml, escapedFont, numbered)
                If Len(rewritten) > 0 And rewritten <> bodyXml Then
                    changed = changed + 1
                    If apply Then
                        If pending = 0 Then
                            skeleton = packageXml
                            skeletonBodyStart = bodyStart
                            skeletonBodyEnd = bodyEnd
                            pendingCharacters = 0
                        End If
                        pending = pending + 1
                        ReDim Preserve starts(1 To pending)
                        ReDim Preserve ends(1 To pending)
                        ReDim Preserve types(1 To pending)
                        ReDim Preserve bodies(1 To pending)
                        starts(pending) = exact.Start
                        ends(pending) = exact.End
                        types(pending) = root.Type
                        bodies(pending) = rewritten
                        pendingCharacters = pendingCharacters + Len(rewritten)
                        If pending >= VT_WORD_MATH_FONT_BATCH_ITEMS Or _
                           pendingCharacters >= VT_WORD_MATH_FONT_BATCH_CHARACTERS Then
                            VTV2CommitMathFontBatch doc, skeleton, skeletonBodyStart, _
                                skeletonBodyEnd, starts, ends, types, bodies, pending
                            pending = 0
                        End If
                    End If
                End If
            End If
        End If
    Next index
    If apply And pending > 0 Then
        VTV2CommitMathFontBatch doc, skeleton, skeletonBodyStart, _
            skeletonBodyEnd, starts, ends, types, bodies, pending
    End If
    VTV2MathFontRepairs = changed
End Function

' Opens the rewritten equations (one paragraph each, in the order given) as
' one hidden document and copies each over the equation it came from.
Private Sub VTV2CommitMathFontBatch( _
    ByVal doc As Document, _
    ByVal skeleton As String, _
    ByVal skeletonBodyStart As Long, _
    ByVal skeletonBodyEnd As Long, _
    ByRef starts() As Long, _
    ByRef ends() As Long, _
    ByRef types() As Long, _
    ByRef bodies() As String, _
    ByVal itemCount As Long)

    Dim stagingPath As String
    Dim staging As Document
    Dim source As Range
    Dim parts() As String
    Dim index As Long
    Dim failureNumber As Long
    Dim failureDescription As String

    ReDim parts(1 To itemCount)
    For index = 1 To itemCount
        parts(index) = bodies(index)
    Next index
    stagingPath = VTApplicationSupportRoot() & "/math-font-" & VTNewUuidV4() & ".xml"
    VTWriteTextAtomic stagingPath, Left$(skeleton, skeletonBodyStart - 1) & _
        Join(parts, "") & Mid$(skeleton, skeletonBodyEnd)
    On Error GoTo Failed
    Set staging = Documents.Open(FileName:=stagingPath, ConfirmConversions:=False, _
        ReadOnly:=True, AddToRecentFiles:=False, Visible:=False)
    If staging.OMaths.Count <> itemCount Then
        Err.Raise vbObjectError + 7615, "VisualTeX", _
            "Word did not open every equation prepared in the new math font."
    End If
    doc.Activate
    For index = 1 To itemCount
        Set source = VTNativeMathRoot(staging.OMaths(index)).Range.Duplicate
        If source.OMaths(1).Type <> types(index) Then source.OMaths(1).Type = types(index)
        VTV2ReplaceEquationWithFont doc, starts(index), ends(index), types(index), source
    Next index
    staging.Close SaveChanges:=wdDoNotSaveChanges
    Set staging = Nothing
    doc.Activate
    On Error Resume Next
    Kill stagingPath
    Err.Clear
    On Error GoTo 0
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Not staging Is Nothing Then staging.Close SaveChanges:=wdDoNotSaveChanges
    doc.Activate
    Kill stagingPath
    On Error GoTo 0
    Err.Raise failureNumber, "VisualTeX equation math font", failureDescription
End Sub

' Copies source (the same equation in the new font) over the equation at
' startPosition..endPosition, keeping its type, alignment and every bookmark
' inside it (formula identity, number targets) at the same offsets.
Private Sub VTV2ReplaceEquationWithFont( _
    ByVal doc As Document, _
    ByVal startPosition As Long, _
    ByVal endPosition As Long, _
    ByVal mathType As Long, _
    ByVal source As Range)

    Dim target As Range
    Dim root As OMath
    Dim mark As Bookmark
    Dim names As New Collection
    Dim offsets As New Collection
    Dim originalText As String
    Dim alignment As Long
    Dim hidden As Boolean
    Dim index As Long

    Set target = doc.Range(startPosition, endPosition)
    If target.OMaths.Count < 1 Then
        Err.Raise vbObjectError + 7615, "VisualTeX", _
            "A Word equation moved while its math font was changed."
    End If
    Set root = VTNativeMathRoot(target.OMaths(1))
    If root.Range.Start <> startPosition Or root.Range.End <> endPosition Then
        Err.Raise vbObjectError + 7615, "VisualTeX", _
            "A Word equation moved while its math font was changed."
    End If
    If mathType = wdOMathDisplay Then alignment = root.Justification
    originalText = target.Text
    hidden = doc.Bookmarks.ShowHidden
    doc.Bookmarks.ShowHidden = True
    For Each mark In target.Bookmarks
        If mark.Start >= startPosition And mark.End <= endPosition Then
            names.Add mark.Name
            offsets.Add Array(mark.Start - startPosition, mark.End - startPosition)
        End If
    Next mark
    doc.Bookmarks.ShowHidden = hidden

    target.FormattedText = source.FormattedText
    Set target = doc.Range(startPosition, endPosition)
    If target.OMaths.Count < 1 Or target.Text <> originalText Then
        Err.Raise vbObjectError + 7615, "VisualTeX", _
            "Word did not keep an equation unchanged when changing its math font."
    End If
    Set root = VTNativeMathRoot(target.OMaths(1))
    If root.Range.Start <> startPosition Or root.Range.End <> endPosition Then
        Err.Raise vbObjectError + 7615, "VisualTeX", _
            "Word did not keep an equation unchanged when changing its math font."
    End If
    If root.Type <> mathType Then root.Type = mathType
    If mathType = wdOMathDisplay Then
        If root.Justification <> alignment Then root.Justification = alignment
    End If
    For index = 1 To names.Count
        If doc.Bookmarks.Exists(CStr(names(index))) Then _
            doc.Bookmarks(CStr(names(index))).Delete
        doc.Bookmarks.Add Name:=CStr(names(index)), Range:=doc.Range( _
            startPosition + offsets(index)(0), startPosition + offsets(index)(1))
    Next index
End Sub

' The body content of a Flat OPC package from Range.WordOpenXML: after
' <w:body> and before its final w:sectPr (or </w:body>).
Private Function VTMathFontPackageBody( _
    ByVal packageXml As String, _
    ByRef bodyStart As Long, _
    ByRef bodyEnd As Long) As Boolean

    Dim closing As Long
    Dim sectionAt As Long

    bodyStart = InStr(1, packageXml, "<w:body>", vbBinaryCompare)
    If bodyStart = 0 Then Exit Function
    bodyStart = bodyStart + Len("<w:body>")
    closing = InStr(bodyStart, packageXml, "</w:body>", vbBinaryCompare)
    If closing = 0 Then Exit Function
    sectionAt = InStrRev(packageXml, "<w:sectPr", closing, vbBinaryCompare)
    If sectionAt >= bodyStart Then
        bodyEnd = sectionAt
    Else
        bodyEnd = closing
    End If
    VTMathFontPackageBody = bodyEnd > bodyStart
End Function

' The start of the next <tagName ...> (exactly that name, so m:r is not
' m:rPr) from position and before limit, or 0.
Private Function VTMathFontNextTag( _
    ByVal xml As String, _
    ByVal tagName As String, _
    ByVal position As Long, _
    ByVal limit As Long) As Long

    Dim found As Long
    Dim following As String

    If position < 1 Then position = 1
    found = InStr(position, xml, "<" & tagName, vbBinaryCompare)
    Do While found > 0 And found < limit
        following = Mid$(xml, found + Len(tagName) + 1, 1)
        If following = ">" Or following = " " Or following = "/" Then
            VTMathFontNextTag = found
            Exit Function
        End If
        found = InStr(found + 1, xml, "<" & tagName, vbBinaryCompare)
    Loop
End Function

' The body with every math run font rewritten (VTV2MathFontRepairs), or ""
' when its equation cannot be read (it is then left as it is).
Private Function VTMathFontRewriteBody( _
    ByVal bodyXml As String, _
    ByVal escapedFont As String, _
    ByVal numbered As Boolean) As String

    Dim mathAt As Long
    Dim regionStart As Long
    Dim regionEnd As Long
    Dim arrayAt As Long
    Dim elementAt As Long
    Dim hashAt As Long
    Dim runAt As Long
    Dim candidate As Long
    Dim textTag As Long

    mathAt = VTMathFontNextTag(bodyXml, "m:oMath", 1, Len(bodyXml) + 1)
    If mathAt = 0 Then Exit Function
    regionStart = InStr(mathAt, bodyXml, ">", vbBinaryCompare) + 1
    regionEnd = InStrRev(bodyXml, "</m:oMath>", -1, vbBinaryCompare)
    If regionStart < 2 Or regionEnd < regionStart Then Exit Function
    If numbered Then
        ' body#(number) is the first row of the outer equation array: change
        ' the body only.
        arrayAt = VTMathFontNextTag(bodyXml, "m:eqArr", regionStart, regionEnd)
        If arrayAt = 0 Then Exit Function
        elementAt = VTMathFontNextTag(bodyXml, "m:e", arrayAt, regionEnd)
        If elementAt = 0 Then Exit Function
        regionStart = InStr(elementAt, bodyXml, ">", vbBinaryCompare) + 1
        hashAt = InStrRev(bodyXml, "#</m:t>", regionEnd, vbBinaryCompare)
        If hashAt < regionStart Then Exit Function
        candidate = VTMathFontNextTag(bodyXml, "m:r", regionStart, hashAt)
        Do While candidate > 0
            runAt = candidate
            candidate = VTMathFontNextTag(bodyXml, "m:r", candidate + 1, hashAt)
        Loop
        If runAt = 0 Then Exit Function
        textTag = InStrRev(bodyXml, "<m:t", hashAt, vbBinaryCompare)
        If textTag > runAt And _
           InStr(textTag, bodyXml, ">", vbBinaryCompare) = hashAt - 1 Then
            ' The # has a run of its own: the body ends before it.
            regionEnd = runAt
        Else
            ' Word merged the # into the body's last run.
            regionEnd = InStr(hashAt, bodyXml, "</m:r>", vbBinaryCompare)
            If regionEnd = 0 Then Exit Function
            regionEnd = regionEnd + Len("</m:r>")
        End If
    End If
    VTMathFontRewriteBody = VTRewriteMathRunFonts( _
        bodyXml, regionStart, regionEnd, escapedFont)
End Function

' Rewrites the fonts of the m:r (not m:nor) and m:ctrlPr elements that start
' in xml(regionStart..regionEnd-1).
Private Function VTRewriteMathRunFonts( _
    ByVal xml As String, _
    ByVal regionStart As Long, _
    ByVal regionEnd As Long, _
    ByVal escapedFont As String) As String

    Dim parts As New Collection
    Dim position As Long
    Dim runAt As Long
    Dim controlAt As Long
    Dim elementAt As Long
    Dim tagEnd As Long
    Dim closeAt As Long
    Dim closing As String
    Dim segment As String
    Dim joined() As String
    Dim index As Long

    parts.Add Left$(xml, regionStart - 1)
    position = regionStart
    Do
        runAt = VTMathFontNextTag(xml, "m:r", position, regionEnd)
        controlAt = VTMathFontNextTag(xml, "m:ctrlPr", position, regionEnd)
        If runAt = 0 And controlAt = 0 Then Exit Do
        If runAt = 0 Or (controlAt > 0 And controlAt < runAt) Then
            elementAt = controlAt
            closing = "</m:ctrlPr>"
        Else
            elementAt = runAt
            closing = "</m:r>"
        End If
        tagEnd = InStr(elementAt, xml, ">", vbBinaryCompare)
        If tagEnd = 0 Then Exit Do
        If Mid$(xml, tagEnd - 1, 1) = "/" Then
            parts.Add Mid$(xml, position, tagEnd + 1 - position)
            position = tagEnd + 1
        Else
            closeAt = InStr(tagEnd, xml, closing, vbBinaryCompare)
            If closeAt = 0 Then Exit Do
            closeAt = closeAt + Len(closing)
            segment = Mid$(xml, elementAt, closeAt - elementAt)
            If closing = "</m:ctrlPr>" Or Not VTMathRunIsNormalText(segment) Then _
                segment = VTRewriteElementFonts(segment, escapedFont)
            parts.Add Mid$(xml, position, elementAt - position)
            parts.Add segment
            position = closeAt
        End If
    Loop
    parts.Add Mid$(xml, position)
    ReDim joined(1 To parts.Count)
    For index = 1 To parts.Count
        joined(index) = parts(index)
    Next index
    VTRewriteMathRunFonts = Join(joined, "")
End Function

' True for an m:r with m:nor (normal text), unless m:val turns it off.
Private Function VTMathRunIsNormalText(ByVal runXml As String) As Boolean
    Dim tagAt As Long
    Dim tagEnd As Long
    Dim tag As String

    tagAt = VTMathFontNextTag(runXml, "m:nor", 1, Len(runXml) + 1)
    If tagAt = 0 Then Exit Function
    tagEnd = InStr(tagAt, runXml, ">", vbBinaryCompare)
    If tagEnd = 0 Then Exit Function
    tag = Mid$(runXml, tagAt, tagEnd - tagAt + 1)
    VTMathRunIsNormalText = InStr(1, tag, "m:val=""0""", vbBinaryCompare) = 0 And _
        InStr(1, tag, "m:val=""off""", vbBinaryCompare) = 0 And _
        InStr(1, tag, "m:val=""false""", vbBinaryCompare) = 0
End Function

' The element with its w:rFonts naming the font in w:ascii and w:hAnsi (theme
' fonts for those two removed). An element without w:rFonts is unchanged.
Private Function VTRewriteElementFonts( _
    ByVal elementXml As String, _
    ByVal escapedFont As String) As String

    Dim fontsAt As Long
    Dim fontsEnd As Long
    Dim tag As String

    VTRewriteElementFonts = elementXml
    fontsAt = VTMathFontNextTag(elementXml, "w:rFonts", 1, Len(elementXml) + 1)
    If fontsAt = 0 Then Exit Function
    fontsEnd = InStr(fontsAt, elementXml, ">", vbBinaryCompare)
    If fontsEnd = 0 Then Exit Function
    tag = Mid$(elementXml, fontsAt, fontsEnd - fontsAt + 1)
    tag = VTXmlTagWithoutAttribute(tag, "w:asciiTheme")
    tag = VTXmlTagWithoutAttribute(tag, "w:hAnsiTheme")
    tag = VTXmlTagWithAttribute(tag, "w:ascii", escapedFont)
    tag = VTXmlTagWithAttribute(tag, "w:hAnsi", escapedFont)
    VTRewriteElementFonts = Left$(elementXml, fontsAt - 1) & tag & _
        Mid$(elementXml, fontsEnd + 1)
End Function

' The position of attributeName in a start tag, or 0.
Private Function VTXmlAttributePosition( _
    ByVal tag As String, ByVal attributeName As String) As Long

    Dim found As Long

    found = InStr(1, tag, " " & attributeName & "=", vbBinaryCompare)
    If found > 0 Then VTXmlAttributePosition = found + 1
End Function

Private Function VTXmlTagWithAttribute( _
    ByVal tag As String, _
    ByVal attributeName As String, _
    ByVal escapedValue As String) As String

    Dim attributeAt As Long
    Dim quoteAt As Long
    Dim valueEnd As Long
    Dim closeAt As Long

    attributeAt = VTXmlAttributePosition(tag, attributeName)
    If attributeAt > 0 Then
        quoteAt = attributeAt + Len(attributeName) + 1
        valueEnd = InStr(quoteAt + 1, tag, Mid$(tag, quoteAt, 1), vbBinaryCompare)
        If valueEnd > 0 Then
            VTXmlTagWithAttribute = Left$(tag, quoteAt) & escapedValue & Mid$(tag, valueEnd)
            Exit Function
        End If
    End If
    If Right$(tag, 2) = "/>" Then
        closeAt = Len(tag) - 1
    Else
        closeAt = Len(tag)
    End If
    VTXmlTagWithAttribute = RTrim$(Left$(tag, closeAt - 1)) & " " & _
        attributeName & "=""" & escapedValue & """" & Mid$(tag, closeAt)
End Function

Private Function VTXmlTagWithoutAttribute( _
    ByVal tag As String, ByVal attributeName As String) As String

    Dim attributeAt As Long
    Dim quoteAt As Long
    Dim valueEnd As Long

    VTXmlTagWithoutAttribute = tag
    attributeAt = VTXmlAttributePosition(tag, attributeName)
    If attributeAt = 0 Then Exit Function
    quoteAt = attributeAt + Len(attributeName) + 1
    valueEnd = InStr(quoteAt + 1, tag, Mid$(tag, quoteAt, 1), vbBinaryCompare)
    If valueEnd = 0 Then Exit Function
    VTXmlTagWithoutAttribute = Left$(tag, attributeAt - 2) & Mid$(tag, valueEnd + 1)
End Function

Private Function VTXmlAttributeValue(ByVal value As String) As String
    value = Replace$(value, "&", "&amp;")
    value = Replace$(value, "<", "&lt;")
    value = Replace$(value, ">", "&gt;")
    VTXmlAttributeValue = Replace$(value, """", "&quot;")
End Function
