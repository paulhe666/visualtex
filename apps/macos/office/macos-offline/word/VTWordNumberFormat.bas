Attribute VB_Name = "VTWordNumberFormat"
Option Explicit
Option Private Module

' The number format: numbering mode and separator, the saved preference, heading scopes and heading numbers.

' Equation numbering style chosen or worked out per document in this Word
' session ("mode|separator", keyed by FullName). Never stored in the document.
Private VT_WORD_SESSION_NUMBERING As Collection
Private VT_WORD_NUMBERING_SCOPE_DEPTH As Long
Private VT_WORD_NUMBERING_SCOPE_CACHED As Boolean
Private VT_WORD_NUMBERING_SCOPE_KEY As String
Private VT_WORD_NUMBERING_SCOPE_MODE As String
Private VT_WORD_NUMBERING_SCOPE_SEPARATOR As String
Private VT_WORD_NUMBERING_PREFERENCE_CACHE_LOADED As Boolean
Private VT_WORD_NUMBERING_PREFERENCE_CACHE_AVAILABLE As Boolean
Private VT_WORD_NUMBERING_PREFERENCE_CACHE_MODE As String
Private VT_WORD_NUMBERING_PREFERENCE_CACHE_SEPARATOR As String

Public Sub VTApplyEquationNumberingFormatPreset(ByVal selectedIndex As Integer)
    Dim numberingMode As String
    Dim separatorText As String
    Dim updated As Long

    On Error GoTo Failed
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 7401, "VisualTeX", _
            "Open a Word document first."
    End If

    Select Case selectedIndex
        Case 0
            numberingMode = VT_WORD_NUMBERING_MODE_SEQUENCE
            separatorText = "."
        Case 1
            numberingMode = VT_WORD_NUMBERING_MODE_CHAPTER
            separatorText = "."
        Case 2
            numberingMode = VT_WORD_NUMBERING_MODE_CHAPTER
            separatorText = "-"
        Case 3
            numberingMode = VT_WORD_NUMBERING_MODE_SECTION
            separatorText = "."
        Case 4
            numberingMode = VT_WORD_NUMBERING_MODE_SECTION
            separatorText = "-"
        Case Else
            Exit Sub
    End Select

    VTSetEquationNumberingFormat _
        ActiveDocument, numberingMode, separatorText
    VTInvalidateWordEquationNumberingFormatControl
    updated = VTUpdateEquationNumbersCore(ActiveDocument, False)
    Application.StatusBar = "VisualTeX: updated " & CStr(updated) & _
        " equation numbers."
    Exit Sub

Failed:
    VTShowError "equation numbering format", Err.Number, Err.Description
End Sub

Public Function VTTryReadEquationNumberingPreference( _
    ByRef numberingMode As String, _
    ByRef separatorText As String) As Boolean

    Dim encodedPreference As String
    Dim preferenceLine As String
    Dim fields As Variant
    Dim storedMode As String
    Dim storedSeparator As String

    numberingMode = ""
    separatorText = ""

    If VT_WORD_NUMBERING_PREFERENCE_CACHE_LOADED Then
        If VT_WORD_NUMBERING_PREFERENCE_CACHE_AVAILABLE Then
            numberingMode = VT_WORD_NUMBERING_PREFERENCE_CACHE_MODE
            separatorText = VT_WORD_NUMBERING_PREFERENCE_CACHE_SEPARATOR
            VTTryReadEquationNumberingPreference = True
        End If
        Exit Function
    End If

    On Error GoTo PreferenceUnavailable
    encodedPreference = VTFileBridgeCall( _
        "ReadVisualTeXNumberingPreference", "")
    If Len(encodedPreference) = 0 Then
        VT_WORD_NUMBERING_PREFERENCE_CACHE_LOADED = True
        VT_WORD_NUMBERING_PREFERENCE_CACHE_AVAILABLE = False
        Exit Function
    End If
    preferenceLine = VTBase64UrlDecodeUtf8(encodedPreference)
    If Len(preferenceLine) = 0 Or Len(preferenceLine) > 256 Then Exit Function

    fields = Split(Trim$(preferenceLine), "|")
    If UBound(fields) <> 1 Then Exit Function
    storedMode = LCase$(Trim$(CStr(fields(0))))
    storedSeparator = Trim$(CStr(fields(1)))
    Select Case storedMode
        Case VT_WORD_NUMBERING_MODE_SEQUENCE
            storedSeparator = "."
        Case VT_WORD_NUMBERING_MODE_CHAPTER, _
             VT_WORD_NUMBERING_MODE_SECTION
            If storedSeparator <> "." And storedSeparator <> "-" Then _
                Exit Function
        Case Else
            Exit Function
    End Select

    numberingMode = storedMode
    separatorText = storedSeparator
    VT_WORD_NUMBERING_PREFERENCE_CACHE_LOADED = True
    VT_WORD_NUMBERING_PREFERENCE_CACHE_AVAILABLE = True
    VT_WORD_NUMBERING_PREFERENCE_CACHE_MODE = storedMode
    VT_WORD_NUMBERING_PREFERENCE_CACHE_SEPARATOR = storedSeparator
    VTTryReadEquationNumberingPreference = True
    Exit Function

PreferenceUnavailable:
    Err.Clear
End Function

Private Sub VTWriteEquationNumberingPreference( _
    ByVal numberingMode As String, _
    ByVal separatorText As String)

    Dim encodedPreference As String
    Dim ignoredResponse As String
    Dim errorNumber As Long
    Dim errorDescription As String

    numberingMode = LCase$(Trim$(numberingMode))
    separatorText = Trim$(separatorText)
    Select Case numberingMode
        Case VT_WORD_NUMBERING_MODE_SEQUENCE
            separatorText = "."
        Case VT_WORD_NUMBERING_MODE_CHAPTER, _
             VT_WORD_NUMBERING_MODE_SECTION
            If separatorText <> "." And separatorText <> "-" Then
                Err.Raise vbObjectError + 7596, "VisualTeX", _
                    "The Equation numbering separator must be . or -."
            End If
        Case Else
            Err.Raise vbObjectError + 7596, "VisualTeX", _
                "The Equation numbering mode is invalid."
    End Select

    On Error GoTo PreferenceWriteFailed
    encodedPreference = VTBase64UrlEncodeUtf8( _
        numberingMode & "|" & separatorText)
    ignoredResponse = VTFileBridgeCall( _
        "WriteVisualTeXNumberingPreference", encodedPreference)
    VT_WORD_NUMBERING_PREFERENCE_CACHE_LOADED = True
    VT_WORD_NUMBERING_PREFERENCE_CACHE_AVAILABLE = True
    VT_WORD_NUMBERING_PREFERENCE_CACHE_MODE = numberingMode
    VT_WORD_NUMBERING_PREFERENCE_CACHE_SEPARATOR = separatorText
    Exit Sub

PreferenceWriteFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    Err.Raise errorNumber, "VisualTeX numbering preference", _
        "Unable to persist the Equation numbering format: " & _
        errorDescription
End Sub

Public Function VTEquationNumberingMode( _
    ByVal documentObject As Document) As String

    Dim numberingMode As String
    Dim separatorText As String

    VTResolveEquationNumbering documentObject, numberingMode, separatorText
    VTEquationNumberingMode = numberingMode
End Function

Public Function VTEquationNumberingSeparator( _
    ByVal documentObject As Document) As String

    Dim numberingMode As String
    Dim separatorText As String

    VTResolveEquationNumbering documentObject, numberingMode, separatorText
    VTEquationNumberingSeparator = separatorText
End Function

' One callback resolves the numbering style of its document dozens of times
' (every chapter prefix, separator and restart check asks again), and an
' answer that has to be worked out walks the document's equations and
' Bookmarks. Inside VisualTeX_ApplyPendingResult the answer cannot change
' except through VTRememberEquationNumbering, which keeps this cache in step.
Public Sub VTBeginNumberingResolveScope()
    If VT_WORD_NUMBERING_SCOPE_DEPTH <= 0 Then
        VT_WORD_NUMBERING_SCOPE_DEPTH = 0
        VT_WORD_NUMBERING_SCOPE_CACHED = False
    End If
    VT_WORD_NUMBERING_SCOPE_DEPTH = VT_WORD_NUMBERING_SCOPE_DEPTH + 1
End Sub

Public Sub VTEndNumberingResolveScope()
    If VT_WORD_NUMBERING_SCOPE_DEPTH > 0 Then _
        VT_WORD_NUMBERING_SCOPE_DEPTH = VT_WORD_NUMBERING_SCOPE_DEPTH - 1
    If VT_WORD_NUMBERING_SCOPE_DEPTH = 0 Then
        VT_WORD_NUMBERING_SCOPE_CACHED = False
        VT_WORD_NUMBERING_SCOPE_KEY = ""
    End If
End Sub

Public Sub VTResolveEquationNumbering( _
    ByVal documentObject As Document, _
    ByRef numberingMode As String, _
    ByRef separatorText As String)

    Dim scopeKey As String

    If VT_WORD_NUMBERING_SCOPE_DEPTH > 0 And Not documentObject Is Nothing Then
        scopeKey = VTSessionNumberingKey(documentObject)
        If VT_WORD_NUMBERING_SCOPE_CACHED And Len(scopeKey) > 0 And _
           StrComp(scopeKey, VT_WORD_NUMBERING_SCOPE_KEY, vbBinaryCompare) = 0 Then
            numberingMode = VT_WORD_NUMBERING_SCOPE_MODE
            separatorText = VT_WORD_NUMBERING_SCOPE_SEPARATOR
            Exit Sub
        End If
    End If
    VTResolveEquationNumberingUncached documentObject, numberingMode, separatorText
    If Len(scopeKey) > 0 Then
        VT_WORD_NUMBERING_SCOPE_KEY = scopeKey
        VT_WORD_NUMBERING_SCOPE_MODE = numberingMode
        VT_WORD_NUMBERING_SCOPE_SEPARATOR = separatorText
        VT_WORD_NUMBERING_SCOPE_CACHED = True
    End If
End Sub

' Word documents stay purely native: the numbering style is never written into
' the document. The style chosen (or worked out) for a document in this Word
' session wins; otherwise the document's own equation numbers show it, and a
' document without numbers uses the user's default. Earlier versions stored it
' in two document variables; they are still honoured when read and removed
' when VisualTeX next numbers the document (VTPinEquationNumbering).
Private Sub VTResolveEquationNumberingUncached( _
    ByVal documentObject As Document, _
    ByRef numberingMode As String, _
    ByRef separatorText As String)

    Dim preferredMode As String
    Dim preferredSeparator As String

    numberingMode = VT_WORD_NUMBERING_MODE_SEQUENCE
    separatorText = "."
    If Not documentObject Is Nothing Then
        If VTTryChosenEquationNumbering( _
           documentObject, numberingMode, separatorText) Then Exit Sub
        If VTTryInferEquationNumbering( _
           documentObject, numberingMode, separatorText) Then
            VTRememberEquationNumbering _
                documentObject, numberingMode, separatorText
            Exit Sub
        End If
    End If
    numberingMode = VT_WORD_NUMBERING_MODE_SEQUENCE
    separatorText = "."
    If VTTryReadEquationNumberingPreference( _
       preferredMode, preferredSeparator) Then
        numberingMode = preferredMode
        separatorText = preferredSeparator
    End If
End Sub

Private Function VTSessionNumberingKey(ByVal documentObject As Document) As String
    On Error Resume Next
    VTSessionNumberingKey = documentObject.FullName
    Err.Clear
    On Error GoTo 0
End Function

Private Sub VTRememberEquationNumbering( _
    ByVal documentObject As Document, _
    ByVal numberingMode As String, _
    ByVal separatorText As String)

    Dim sessionKey As String

    sessionKey = VTSessionNumberingKey(documentObject)
    If Len(sessionKey) = 0 Then Exit Sub
    If VT_WORD_SESSION_NUMBERING Is Nothing Then _
        Set VT_WORD_SESSION_NUMBERING = New Collection
    On Error Resume Next
    VT_WORD_SESSION_NUMBERING.Remove sessionKey
    Err.Clear
    On Error GoTo 0
    VT_WORD_SESSION_NUMBERING.Add numberingMode & "|" & separatorText, sessionKey
    If VT_WORD_NUMBERING_SCOPE_DEPTH > 0 Then
        VT_WORD_NUMBERING_SCOPE_KEY = sessionKey
        VT_WORD_NUMBERING_SCOPE_MODE = numberingMode
        VT_WORD_NUMBERING_SCOPE_SEPARATOR = separatorText
        VT_WORD_NUMBERING_SCOPE_CACHED = True
    End If
End Sub

' The style chosen for this document (this session, or an earlier version's
' document variables), without looking at its numbers.
Private Function VTTryChosenEquationNumbering( _
    ByVal documentObject As Document, _
    ByRef numberingMode As String, _
    ByRef separatorText As String) As Boolean

    Dim sessionKey As String
    Dim chosen As String
    Dim parts() As String
    Dim storedMode As String
    Dim storedSeparator As String

    sessionKey = VTSessionNumberingKey(documentObject)
    If Len(sessionKey) > 0 And Not VT_WORD_SESSION_NUMBERING Is Nothing Then
        On Error Resume Next
        chosen = VT_WORD_SESSION_NUMBERING(sessionKey)
        Err.Clear
        On Error GoTo 0
        If Len(chosen) > 0 Then
            parts = Split(chosen, "|")
            numberingMode = parts(0)
            separatorText = parts(1)
            VTTryChosenEquationNumbering = True
            Exit Function
        End If
    End If
    If VTTryGetDocumentVariable( _
       documentObject, VT_WORD_NUMBERING_MODE_VARIABLE, storedMode) Then
        storedMode = LCase$(Trim$(storedMode))
        Select Case storedMode
            Case VT_WORD_NUMBERING_MODE_SEQUENCE, _
                 VT_WORD_NUMBERING_MODE_CHAPTER, _
                 VT_WORD_NUMBERING_MODE_SECTION
                numberingMode = storedMode
                separatorText = "."
                If VTTryGetDocumentVariable( _
                   documentObject, VT_WORD_NUMBERING_SEPARATOR_VARIABLE, _
                   storedSeparator) Then
                    If Trim$(storedSeparator) = "-" Then separatorText = "-"
                End If
                If numberingMode = VT_WORD_NUMBERING_MODE_SEQUENCE Then _
                    separatorText = "."
                VTTryChosenEquationNumbering = True
        End Select
    End If
End Function

' The style the first equation number in the document shows: a Word equation
' (#(...) with its SEQ field) or a numbered image formula, whichever comes
' first. Read from the visible prefix, so numbers made on Windows count too:
' none -> sequence, "2." -> chapter, "2.1-" -> section with a dash.
Private Function VTTryInferEquationNumbering( _
    ByVal documentObject As Document, _
    ByRef numberingMode As String, _
    ByRef separatorText As String) As Boolean

    Dim math As OMath
    Dim numberField As Field
    Dim exact As Range
    Dim marker As Range
    Dim candidate As Bookmark
    Dim index As Long
    Dim firstStart As Long
    Dim firstPrefix As String
    Dim found As Boolean
    Dim numberText As String
    Dim ordinalStart As Long
    Dim roots As Collection
    Dim searchFrom As Long
    Dim pictureField As Field
    Dim picturePrefix As String

    On Error GoTo InferFailed
    ' The first numbered equation decides. Its root is asked for a few at a time: a collection of every root of a long document
    ' costs 0.3 ms each (100 ms for 400 equations) and only the first one is read.
    searchFrom = documentObject.Content.Start
    Do
        Set roots = VTNumberFieldNativeMathRoots( _
            documentObject, searchFrom, documentObject.Content.End, False, 4)
        If roots.Count = 0 Then Exit Do
        For Each exact In roots
            searchFrom = exact.End
            Set numberField = VTPureNativeNumberField(exact)
            If Not numberField Is Nothing Then
                Set marker = VTNativeEquationArrayMarkerRange(exact)
                If Not marker Is Nothing Then
                    firstStart = exact.Start
                    firstPrefix = documentObject.Range( _
                        marker.End, VTEquationFieldStart(numberField)).Text
                    found = True
                    Exit Do
                End If
            End If
        Next exact
        If roots.Count < 4 Then Exit Do
    Loop
    For Each candidate In documentObject.Bookmarks
        If Left$(candidate.Name, Len(VT_WORD_NUMBER_BOOKMARK_PREFIX)) = _
           VT_WORD_NUMBER_BOOKMARK_PREFIX Then
            If Not found Or candidate.Range.Start < firstStart Then
                numberText = Trim$(candidate.Range.Text)
                ordinalStart = Len(numberText)
                Do While ordinalStart > 0
                    If Mid$(numberText, ordinalStart, 1) Like "[0-9]" Then
                        ordinalStart = ordinalStart - 1
                    Else
                        Exit Do
                    End If
                Loop
                If ordinalStart < Len(numberText) Then
                    firstStart = candidate.Range.Start
                    firstPrefix = Left$(numberText, ordinalStart)
                    found = True
                End If
            End If
        End If
    Next candidate
    ' A numbered picture: the prefix written in its VisualTeXPlaceRef. Its
    ' number bookmark spans the whole field, whose text does not show it.
    For Each pictureField In documentObject.Fields
        If pictureField.Type = wdFieldMacroButton Then
            If VTIsImagePlaceRef(pictureField) Then
                If Not found Or VTEquationFieldStart(pictureField) < firstStart Then
                    If VTTryImageNumberPrefixText(pictureField, picturePrefix) Then
                        firstStart = VTEquationFieldStart(pictureField)
                        firstPrefix = picturePrefix
                        found = True
                    End If
                End If
                Exit For
            End If
        End If
    Next pictureField
    If Not found Then Exit Function
    VTTryInferEquationNumbering = VTParseEquationNumberPrefix( _
        firstPrefix, numberingMode, separatorText)
    Exit Function

InferFailed:
    VTTryInferEquationNumbering = False
End Function

Private Function VTParseEquationNumberPrefix( _
    ByVal prefixText As String, _
    ByRef numberingMode As String, _
    ByRef separatorText As String) As Boolean

    Dim cleaned As String
    Dim body As String
    Dim position As Long

    For position = 1 To Len(prefixText)
        Select Case AscW(Mid$(prefixText, position, 1))
            Case 48 To 57, 46
                cleaned = cleaned & Mid$(prefixText, position, 1)
            Case 45, 8208, 8209, 8211, 8722
                cleaned = cleaned & "-"
        End Select
    Next position
    If Len(cleaned) = 0 Then
        numberingMode = VT_WORD_NUMBERING_MODE_SEQUENCE
        separatorText = "."
        VTParseEquationNumberPrefix = True
        Exit Function
    End If
    separatorText = Right$(cleaned, 1)
    If separatorText <> "." And separatorText <> "-" Then Exit Function
    body = Left$(cleaned, Len(cleaned) - 1)
    If Len(body) = 0 Or InStr(1, body, "-") > 0 Then Exit Function
    If InStr(1, body, ".") = 0 Then
        numberingMode = VT_WORD_NUMBERING_MODE_CHAPTER
    Else
        numberingMode = VT_WORD_NUMBERING_MODE_SECTION
    End If
    VTParseEquationNumberPrefix = True
End Function

' Numbering the document now: keep its style for the session and remove an
' earlier version's document variables. Only called while VisualTeX changes
' numbers, so an untouched document is never marked as changed.
Public Sub VTPinEquationNumbering(ByVal documentObject As Document)
    Dim numberingMode As String
    Dim separatorText As String

    If documentObject Is Nothing Then Exit Sub
    VTResolveEquationNumbering documentObject, numberingMode, separatorText
    VTRememberEquationNumbering documentObject, numberingMode, separatorText
    VTDeleteDocumentVariable documentObject, VT_WORD_NUMBERING_MODE_VARIABLE
    VTDeleteDocumentVariable documentObject, VT_WORD_NUMBERING_SEPARATOR_VARIABLE
End Sub

Public Function VTEquationNumberingDisplaySeparator( _
    ByVal documentObject As Document) As String

    If VTEquationNumberingSeparator(documentObject) = "-" Then
        ' Keep chapter/section Equation numbers typographically compact. Word
        ' treats U+2212 MINUS SIGN (and an ASCII hyphen promoted by OMath) as a
        ' binary math operator, which adds visible spacing around the separator
        ' inside numbered OMML. U+2010 HYPHEN is a punctuation separator instead:
        ' the shared VT_N_ Bookmark therefore renders identically in image text,
        ' native OMath REF results and ordinary body cross-references.
        VTEquationNumberingDisplaySeparator = ChrW(8208)
    Else
        VTEquationNumberingDisplaySeparator = "."
    End If
End Function

Public Function VTComparableEquationNumberText( _
    ByVal numberText As String) As String

    numberText = Trim$(numberText)
    numberText = Replace$(numberText, ChrW(8722), "-")
    numberText = Replace$(numberText, ChrW(8208), "-")
    numberText = Replace$(numberText, ChrW(8209), "-")
    VTComparableEquationNumberText = numberText
End Function

Public Sub VTSetEquationNumberingFormat( _
    ByVal documentObject As Document, _
    ByVal numberingMode As String, _
    ByVal separatorText As String)

    numberingMode = LCase$(Trim$(numberingMode))
    If documentObject Is Nothing Then
        Err.Raise vbObjectError + 7596, "VisualTeX", _
            "The Equation numbering format requires an open document."
    End If
    Select Case numberingMode
        Case VT_WORD_NUMBERING_MODE_SEQUENCE
            separatorText = "."
        Case VT_WORD_NUMBERING_MODE_CHAPTER, _
             VT_WORD_NUMBERING_MODE_SECTION
            If separatorText <> "." And separatorText <> "-" Then
                Err.Raise vbObjectError + 7596, "VisualTeX", _
                    "The Equation numbering separator must be . or -."
            End If
        Case Else
            Err.Raise vbObjectError + 7596, "VisualTeX", _
                "The Equation numbering mode is invalid."
    End Select
    VTWriteEquationNumberingPreference numberingMode, separatorText
    VTRememberEquationNumbering documentObject, numberingMode, separatorText
    VTDeleteDocumentVariable documentObject, VT_WORD_NUMBERING_MODE_VARIABLE
    VTDeleteDocumentVariable documentObject, VT_WORD_NUMBERING_SEPARATOR_VARIABLE
End Sub

Public Function VTEquationNumberingRestartLevel( _
    ByVal documentObject As Document) As Long

    Select Case VTEquationNumberingMode(documentObject)
        Case VT_WORD_NUMBERING_MODE_CHAPTER
            VTEquationNumberingRestartLevel = 1
        Case VT_WORD_NUMBERING_MODE_SECTION
            VTEquationNumberingRestartLevel = 2
    End Select
End Function

' Heading paragraphs with an outline level from 1 to maxLevel that start before
' searchEnd, in document order, as Array(start, level). Word's Find visits only
' the paragraphs with the requested outline level (about 4 ms per hit on Word
' for Mac), while reading OutlineLevel paragraph by paragraph costs 0.2-0.3 ms
' for every paragraph of a long document. Find cannot match the end-of-cell
' mark, so with includeTables the paragraphs of tables are read directly and
' the result equals a walk over every paragraph.
Private Function VTOutlineHeadingParagraphs( _
    ByVal documentObject As Document, _
    ByVal maxLevel As Long, _
    ByVal searchEnd As Long, _
    ByVal includeTables As Boolean, _
    Optional ByVal searchStart As Long = -1) As Collection

    Dim result As Collection
    Dim scanRange As Range
    Dim paragraphRange As Range
    Dim tableObject As Table
    Dim tableParagraph As Paragraph
    Dim level As Long
    Dim headingStart As Long
    Dim previousEnd As Long
    Dim tableLevel As Long

    Set result = New Collection
    Set VTOutlineHeadingParagraphs = result
    If documentObject Is Nothing Or maxLevel < 1 Then Exit Function
    If searchEnd > documentObject.Content.End Then _
        searchEnd = documentObject.Content.End
    If searchStart < documentObject.Content.Start Then _
        searchStart = documentObject.Content.Start
    If searchEnd <= searchStart Then Exit Function
    For level = 1 To maxLevel
        Set scanRange = documentObject.Range( _
            Start:=searchStart, End:=documentObject.Content.End)
        With scanRange.Find
            .ClearFormatting
            .Text = "^p"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .ParagraphFormat.OutlineLevel = level
        End With
        previousEnd = -1
        Do While scanRange.Find.Execute
            If scanRange.End <= previousEnd Then Exit Do
            previousEnd = scanRange.End
            Set paragraphRange = scanRange.Paragraphs(1).Range
            headingStart = paragraphRange.Start
            If headingStart >= searchEnd Then Exit Do
            If headingStart >= searchStart Then _
                VTAddHeadingParagraphEntry result, headingStart, level
            scanRange.Collapse wdCollapseEnd
        Loop
        scanRange.Find.ClearFormatting
    Next level
    If includeTables And searchEnd > documentObject.Content.Start Then
        For Each tableObject In documentObject.Range( _
            Start:=searchStart, End:=searchEnd).Tables
            For Each tableParagraph In tableObject.Range.Paragraphs
                headingStart = tableParagraph.Range.Start
                If headingStart >= searchEnd Then Exit For
                If headingStart < searchStart Then GoTo NextTableParagraph
                tableLevel = tableParagraph.OutlineLevel
                If tableLevel >= 1 And tableLevel <= maxLevel Then _
                    VTAddHeadingParagraphEntry result, headingStart, tableLevel
NextTableParagraph:
            Next tableParagraph
        Next tableObject
    End If
End Function

' Inserts Array(start, level) into entries ordered by start, once per start.
Private Sub VTAddHeadingParagraphEntry( _
    ByVal entries As Collection, _
    ByVal headingStart As Long, _
    ByVal level As Long)

    Dim index As Long

    For index = entries.Count To 1 Step -1
        If entries(index)(0) = headingStart Then Exit Sub
        If entries(index)(0) < headingStart Then
            entries.Add Array(headingStart, level), After:=index
            Exit Sub
        End If
    Next index
    If entries.Count = 0 Then
        entries.Add Array(headingStart, level)
    Else
        entries.Add Array(headingStart, level), Before:=1
    End If
End Sub

' The chapter ("3") or section ("3.2") number of the formula at formulaRange,
' by the rule of the OMML numbers (VTV2HeadingScope): the number written in the
' last heading of that level before the formula, from Word's list numbering or
' the heading text; headings are counted only in a document whose headings
' carry no number. "0" / "0.0" before any heading.
Public Function VTEquationHeadingPrefix( _
    ByVal formulaRange As Range, _
    ByVal numberingMode As String) As String

    Dim targetLevel As Long
    Dim scopeStart As Long
    Dim automatic As Boolean

    If formulaRange Is Nothing Then Exit Function
    numberingMode = LCase$(Trim$(numberingMode))
    If numberingMode = VT_WORD_NUMBERING_MODE_CHAPTER Then
        targetLevel = 1
    ElseIf numberingMode = VT_WORD_NUMBERING_MODE_SECTION Then
        targetLevel = 2
    Else
        Exit Function
    End If
    VTEquationHeadingPrefix = VTV2HeadingScope( _
        VTV2PlanHeadings(formulaRange.Document, targetLevel, formulaRange.Start), _
        formulaRange.Start, targetLevel, scopeStart, automatic)
End Function

Public Sub VTBuildEquationHeadingPrefixes( _
    ByVal documentObject As Document, _
    ByRef sequenceAnchors() As Long, _
    ByVal sequenceCount As Long, _
    ByVal numberingMode As String, _
    ByRef headingPrefixes() As String, _
    ByRef localOrdinals() As Long)

    Dim candidateParagraph As Paragraph
    Dim candidateLevel As Long
    Dim targetLevel As Long
    Dim itemIndex As Long
    Dim paragraphStart As Long
    Dim currentLocalOrdinal As Long
    Dim headings As Collection
    Dim scopeStart As Long
    Dim automatic As Boolean

    If documentObject Is Nothing Or sequenceCount <= 0 Then Exit Sub
    numberingMode = LCase$(Trim$(numberingMode))
    If numberingMode = VT_WORD_NUMBERING_MODE_CHAPTER Then
        targetLevel = 1
    ElseIf numberingMode = VT_WORD_NUMBERING_MODE_SECTION Then
        targetLevel = 2
    Else
        Exit Sub
    End If

    ReDim headingPrefixes(1 To sequenceCount)
    ReDim localOrdinals(1 To sequenceCount)

    ' Chapter and section numbers follow the OMML numbers (VTV2HeadingScope).
    Set headings = VTV2CollectHeadings(documentObject, targetLevel)
    For itemIndex = 1 To sequenceCount
        headingPrefixes(itemIndex) = VTV2HeadingScope( _
            headings, sequenceAnchors(itemIndex), targetLevel, _
            scopeStart, automatic)
    Next itemIndex

    ' The running number restarts at every heading of the numbering level, as
    ' SEQ \s does for OMML numbers. sequenceAnchors are sorted, so one forward
    ' paragraph pass assigns the local ordinals.
    itemIndex = 1
    For Each candidateParagraph In documentObject.Paragraphs
        paragraphStart = candidateParagraph.Range.Start
        Do While itemIndex <= sequenceCount
            If sequenceAnchors(itemIndex) > paragraphStart Then Exit Do
            currentLocalOrdinal = currentLocalOrdinal + 1
            localOrdinals(itemIndex) = currentLocalOrdinal
            itemIndex = itemIndex + 1
        Loop
        If itemIndex > sequenceCount Then Exit For
        candidateLevel = candidateParagraph.OutlineLevel
        If candidateLevel >= 1 And candidateLevel <= targetLevel Then
            currentLocalOrdinal = 0
        End If
    Next candidateParagraph
    Do While itemIndex <= sequenceCount
        currentLocalOrdinal = currentLocalOrdinal + 1
        localOrdinals(itemIndex) = currentLocalOrdinal
        itemIndex = itemIndex + 1
    Loop
End Sub

' --- small text helpers (Mac VBA has no regular expressions) ----------------

Private Function VTV2CodeUnit(ByVal character As String) As Long
    VTV2CodeUnit = AscW(character)
    If VTV2CodeUnit < 0 Then VTV2CodeUnit = VTV2CodeUnit + 65536
End Function

Private Function VTV2IsDigit(ByVal character As String) As Boolean
    Dim code As Long
    If Len(character) = 0 Then Exit Function
    code = VTV2CodeUnit(character)
    ' .NET \d (Windows) accepts ASCII and fullwidth digits in these labels.
    VTV2IsDigit = (code >= 48 And code <= 57) Or _
        (code >= &HFF10& And code <= &HFF19&)
End Function

Public Function VTV2IsSpace(ByVal character As String) As Boolean
    Dim code As Long
    If Len(character) = 0 Then Exit Function
    code = VTV2CodeUnit(character)
    ' .NET char.IsWhiteSpace.
    Select Case code
        Case 9 To 13, 32, 133, 160, 5760, 8192 To 8202, 8232, 8233, 8239, 8287, 12288
            VTV2IsSpace = True
    End Select
End Function

Public Function VTV2TrimSpaces(ByVal value As String) As String
    Do While Len(value) > 0
        If Not VTV2IsSpace(Left$(value, 1)) Then Exit Do
        value = Mid$(value, 2)
    Loop
    Do While Len(value) > 0
        If Not VTV2IsSpace(Right$(value, 1)) Then Exit Do
        value = Left$(value, Len(value) - 1)
    Loop
    VTV2TrimSpaces = value
End Function

Private Function VTV2IsWordCharacter(ByVal character As String) As Boolean
    Dim code As Long
    If Len(character) = 0 Then Exit Function
    code = VTV2CodeUnit(character)
    VTV2IsWordCharacter = (code >= 48 And code <= 57) Or _
        (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or _
        code = 95 Or code > 127
End Function

Private Function VTV2RemoveWholeWord( _
    ByVal text As String, ByVal word As String) As String

    Dim lowered As String
    Dim position As Long
    Dim before As String
    Dim after As String

    Do
        lowered = LCase$(text)
        position = InStr(1, lowered, word, vbBinaryCompare)
        Do While position > 0
            before = ""
            If position > 1 Then before = Mid$(text, position - 1, 1)
            after = Mid$(text, position + Len(word), 1)
            If Not VTV2IsWordCharacter(before) And _
               Not VTV2IsWordCharacter(after) Then Exit Do
            position = InStr(position + 1, lowered, word, vbBinaryCompare)
        Loop
        If position = 0 Then Exit Do
        text = Left$(text, position - 1) & Mid$(text, position + Len(word))
    Loop
    VTV2RemoveWholeWord = text
End Function

' Text sets of Windows' heading-number rules, spelled with ChrW so this module
' stays ASCII like the rest of the add-in source.
Private Function VTV2HeadingDecorations() As String
    ' The CJK chapter words U+7B2C U+7AE0 U+8282 U+7BC7 U+90E8 U+5377 and
    ' ASCII/fullwidth punctuation, dashes and brackets.
    VTV2HeadingDecorations = ChrW(&H7B2C&) & ChrW(&H7AE0&) & ChrW(&H8282&) & _
        ChrW(&H7BC7&) & ChrW(&H90E8&) & ChrW(&H5377&) & "." & ChrW(&HFF0E&) & _
        ChrW(&H3001&) & ",:" & ChrW(&HFF1A&) & ";" & ChrW(&HFF1B&) & "-" & _
        ChrW(&H2013&) & ChrW(&H2014&) & "()" & ChrW(&HFF08&) & ChrW(&HFF09&) & _
        "[]" & ChrW(&H3010&) & ChrW(&H3011&) & "{}"
End Function

Private Function VTV2HeadingNumberSeparators() As String
    ' Full stop, fullwidth full stop, hyphen, en dash, em dash.
    VTV2HeadingNumberSeparators = "." & ChrW(&HFF0E&) & "-" & ChrW(&H2013&) & ChrW(&H2014&)
End Function

Private Function VTV2HeadingNumberTerminators() As String
    ' U+7AE0 U+8282 U+7BC7 U+90E8, full stops, ideographic comma, colons,
    ' closing parentheses and dashes.
    VTV2HeadingNumberTerminators = ChrW(&H7AE0&) & ChrW(&H8282&) & ChrW(&H7BC7&) & _
        ChrW(&H90E8&) & "." & ChrW(&HFF0E&) & ChrW(&H3001&) & ":" & ChrW(&HFF1A&) & _
        ")" & ChrW(&HFF09&) & "-" & ChrW(&H2013&) & ChrW(&H2014&)
End Function

' Windows WordEquationNumbering.NormalizeHeadingNumberText: a list label such
' such as "3.2", "Chapter 3" or the CJK "chapter 3" form becomes "3" / "3.2";
' any other text fails.
Private Function VTV2HeadingListNumber( _
    ByVal value As String, ByVal outlineLevel As Long) As String

    Dim text As String
    Dim remainder As String
    Dim numbers As String
    Dim current As String
    Dim character As String
    Dim word As Variant
    Dim count As Long
    Dim index As Long

    text = Replace$(Replace$(Replace$(Replace$(value, vbTab, ""), vbCr, ""), vbLf, ""), Chr$(7), "")
    text = VTV2TrimSpaces(text)
    If Len(text) = 0 Then Exit Function
    For index = 1 To Len(text)
        character = Mid$(text, index, 1)
        If VTV2IsDigit(character) Then
            If VTV2CodeUnit(character) >= &HFF10& Then _
                character = ChrW(VTV2CodeUnit(character) - &HFF10& + 48)
            current = current & character
        Else
            If Len(current) > 0 Then
                numbers = numbers & IIf(Len(numbers) > 0, ".", "") & current
                count = count + 1
                current = ""
            End If
            remainder = remainder & character
        End If
    Next index
    If Len(current) > 0 Then
        numbers = numbers & IIf(Len(numbers) > 0, ".", "") & current
        count = count + 1
    End If
    If count = 0 Or count > IIf(outlineLevel > 1, outlineLevel, 1) Then Exit Function

    ' (?i:\b(?:chapter|section|part)\b) after the digits are gone.
    For Each word In Array("chapter", "section", "part")
        remainder = VTV2RemoveWholeWord(remainder, CStr(word))
    Next word
    For index = Len(remainder) To 1 Step -1
        character = Mid$(remainder, index, 1)
        If InStr(1, VTV2HeadingDecorations(), character, vbBinaryCompare) > 0 Or _
           VTV2IsSpace(character) Then
            remainder = Left$(remainder, index - 1) & Mid$(remainder, index + 1)
        End If
    Next index
    If Len(remainder) > 0 Then Exit Function
    VTV2HeadingListNumber = numbers
End Function

' Windows WordEquationNumbering.ParseHeadingNumberFromText: typed headings such
' such as "8. Title", "8-1 Title" or the CJK "chapter 8" form.
Private Function VTV2HeadingTextNumber( _
    ByVal value As String, ByVal outlineLevel As Long) As String

    Dim text As String
    Dim position As Long
    Dim afterNumber As Long
    Dim digitsEnd As Long
    Dim probe As Long
    Dim spaces As Long
    Dim character As String
    Dim number As String
    Dim component As String
    Dim boundaries As New Collection
    Dim numbers As New Collection
    Dim index As Long
    Dim parts() As String
    Dim count As Long

    text = Replace$(Replace$(Replace$(value, vbCr, ""), vbLf, ""), Chr$(7), "")
    text = VTV2TrimSpaces(text)
    If Len(text) = 0 Then Exit Function
    position = 1
    If Mid$(text, 1, 1) = ChrW(&H7B2C&) Then
        position = 2
        Do While position <= Len(text)
            If Not VTV2IsSpace(Mid$(text, position, 1)) Then Exit Do
            position = position + 1
        Loop
    End If
    ' (?<number>\d+(?:\s*[separator]\s*\d+)*): collect every component boundary
    ' so the regex backtracking can be replayed from the longest match down.
    digitsEnd = position
    Do While digitsEnd <= Len(text)
        If Not VTV2IsDigit(Mid$(text, digitsEnd, 1)) Then Exit Do
        digitsEnd = digitsEnd + 1
    Loop
    If digitsEnd = position Then Exit Function
    number = Mid$(text, position, digitsEnd - position)
    boundaries.Add digitsEnd
    numbers.Add number
    probe = digitsEnd
    Do
        afterNumber = probe
        Do While afterNumber <= Len(text)
            If Not VTV2IsSpace(Mid$(text, afterNumber, 1)) Then Exit Do
            afterNumber = afterNumber + 1
        Loop
        If afterNumber > Len(text) Then Exit Do
        If InStr(1, VTV2HeadingNumberSeparators(), Mid$(text, afterNumber, 1), vbBinaryCompare) = 0 Then Exit Do
        afterNumber = afterNumber + 1
        Do While afterNumber <= Len(text)
            If Not VTV2IsSpace(Mid$(text, afterNumber, 1)) Then Exit Do
            afterNumber = afterNumber + 1
        Loop
        digitsEnd = afterNumber
        Do While digitsEnd <= Len(text)
            If Not VTV2IsDigit(Mid$(text, digitsEnd, 1)) Then Exit Do
            digitsEnd = digitsEnd + 1
        Loop
        If digitsEnd = afterNumber Then Exit Do
        component = Mid$(text, afterNumber, digitsEnd - afterNumber)
        number = number & "." & component
        boundaries.Add digitsEnd
        numbers.Add number
        probe = digitsEnd
    Loop
    ' (?:\s*[CJK unit]|\s*[terminator]|\s+) after the number.
    For index = boundaries.Count To 1 Step -1
        probe = boundaries(index)
        spaces = 0
        Do While probe <= Len(text)
            If Not VTV2IsSpace(Mid$(text, probe, 1)) Then Exit Do
            probe = probe + 1
            spaces = spaces + 1
        Loop
        character = ""
        If probe <= Len(text) Then character = Mid$(text, probe, 1)
        If (Len(character) > 0 And InStr(1, VTV2HeadingNumberTerminators(), character, vbBinaryCompare) > 0) Or _
           spaces > 0 Then
            parts = Split(numbers(index), ".")
            count = UBound(parts) + 1
            If count = 0 Or count > IIf(outlineLevel > 1, outlineLevel, 1) Then Exit Function
            VTV2HeadingTextNumber = numbers(index)
            Exit Function
        End If
    Next index
End Function

' --- heading scope (Windows ResolveHeadingScopeAtPosition) -------------------

' Every heading paragraph of outline level 1..targetLevel outside tables and
' frames, in document order (Windows heading candidates): each item is
' Array(start, outline level, explicit number, has Word list numbering).
Public Function VTV2CollectHeadings( _
    ByVal doc As Document, ByVal targetLevel As Long) As Collection

    Set VTV2CollectHeadings = VTV2CollectHeadingsBetween( _
        doc, targetLevel, doc.Content.Start, doc.Content.End)
End Function

' The headings of VTV2CollectHeadings that start at or after rangeStart and
' before rangeEnd.
Private Function VTV2CollectHeadingsBetween( _
    ByVal doc As Document, _
    ByVal targetLevel As Long, _
    ByVal rangeStart As Long, _
    ByVal rangeEnd As Long) As Collection

    Dim result As New Collection
    Dim headingEntry As Variant
    Dim paragraphRange As Range
    Dim outlineLevel As Long
    Dim explicitNumber As String

    If targetLevel >= 1 Then
        ' Only headings matter; visit just them (VTOutlineHeadingParagraphs).
        ' Headings inside tables are excluded below, so Find alone suffices.
        For Each headingEntry In VTOutlineHeadingParagraphs( _
            doc, targetLevel, rangeEnd, False, rangeStart)
            outlineLevel = headingEntry(1)
            If outlineLevel >= 1 And outlineLevel <= targetLevel Then
                Set paragraphRange = doc.Range( _
                    Start:=headingEntry(0), End:=headingEntry(0)).Paragraphs(1).Range
                If Not VTRangeIsWithinTable(paragraphRange) Then
                    If paragraphRange.Frames.Count = 0 Then
                        explicitNumber = ""
                        On Error Resume Next
                        explicitNumber = VTV2HeadingListNumber( _
                            paragraphRange.ListFormat.ListString, outlineLevel)
                        On Error GoTo 0
                        If Len(explicitNumber) = 0 Then
                            explicitNumber = VTV2HeadingTextNumber( _
                                paragraphRange.Text, outlineLevel)
                        End If
                        result.Add Array(paragraphRange.Start, outlineLevel, explicitNumber, _
                            paragraphRange.ListFormat.ListType <> wdListNoNumbering)
                    End If
                End If
            End If
        Next headingEntry
    End If
    Set VTV2CollectHeadingsBetween = result
End Function

' The chapter/section number of the formula at position for a heading level of
' 1 or 2 (Windows ResolveHeadingScopeAtPosition): "3", "3.2", or "0" / "0.0"
' before any heading. scopeStart is the heading paragraph's start, or -1 when
' no heading applies. automatic is True when that heading carries Word list
' numbering; its prefix is then read live with STYLEREF, otherwise written as
' text (Windows ResolveLiteralHeadingPrefix).
Public Function VTV2HeadingScope( _
    ByVal headings As Collection, _
    ByVal position As Long, _
    ByVal targetLevel As Long, _
    ByRef scopeStart As Long, _
    ByRef automatic As Boolean) As String

    Dim usesExplicit As Boolean
    Dim index As Long
    Dim level As Long
    Dim chosen As Long
    Dim counters(1 To 9) As Long
    Dim deeper As Long
    Dim numberText As String
    Dim anchorNumber As String
    Dim anchorIndex As Long

    scopeStart = -1
    automatic = False
    If targetLevel < 1 Then Exit Function

    ' Explicit scope (Windows TryResolveExplicitHeadingScopeAtPosition): the
    ' last numbered heading of each level up to the formula, each deeper level
    ' only inside the scope chosen for the level above it.
    For level = 1 To targetLevel
        For index = headings.Count To 1 Step -1
            If headings(index)(1) = level And headings(index)(0) <= position Then
                If chosen > 0 Then
                    If headings(index)(0) <= headings(chosen)(0) Then Exit For
                End If
                If Len(headings(index)(2)) > 0 Then
                    chosen = index
                    numberText = headings(index)(2)
                    If level < targetLevel Then _
                        numberText = numberText & Replace$(String$(targetLevel - level, "x"), "x", ".0")
                    Exit For
                End If
            End If
        Next index
    Next level
    If chosen > 0 Then
        scopeStart = headings(chosen)(0)
        automatic = headings(chosen)(3)
        VTV2HeadingScope = numberText
        Exit Function
    End If

    ' Fallback (Windows GetHeadingNumberAnchors + ResolveEquationNumberScope):
    ' count unnumbered headings only in a document without explicit numbers.
    For index = 1 To headings.Count
        If Len(headings(index)(2)) > 0 Then usesExplicit = True
    Next index
    For index = 1 To headings.Count
        numberText = headings(index)(2)
        If Len(numberText) = 0 Then
            If usesExplicit Then GoTo NextAnchor
            level = headings(index)(1)
            counters(level) = counters(level) + 1
            For deeper = level + 1 To 9
                counters(deeper) = 0
            Next deeper
            For deeper = 1 To level
                numberText = numberText & IIf(deeper > 1, ".", "") & _
                    CStr(IIf(counters(deeper) > 1, counters(deeper), 1))
            Next deeper
        End If
        If headings(index)(1) < targetLevel Then _
            numberText = numberText & Replace$(String$(targetLevel - headings(index)(1), "x"), "x", ".0")
        If headings(index)(0) > position Then Exit For
        anchorIndex = index
        anchorNumber = numberText
NextAnchor:
    Next index
    If anchorIndex > 0 Then
        scopeStart = headings(anchorIndex)(0)
        automatic = headings(anchorIndex)(3)
        VTV2HeadingScope = anchorNumber
    Else
        VTV2HeadingScope = IIf(targetLevel = 1, "0", "0.0")
    End If
End Function

' The headings VTV2PlanInsertion reads (VTV2CollectHeadings). During a
' document import the headings before the paragraph that held the previous
' position come from the cache; the ones added since and every heading from
' the current paragraph on are read again.
Public Function VTV2PlanHeadings( _
    ByVal doc As Document, _
    ByVal targetLevel As Long, _
    ByVal position As Long) As Collection

    Dim cached As Collection
    Dim result As Collection
    Dim entry As Variant
    Dim boundary As Long
    Dim usable As Boolean

    If Not VT_WORD_V2_IMPORT_CACHE_ACTIVE Or targetLevel < 1 Then
        Set VTV2PlanHeadings = VTV2CollectHeadings(doc, targetLevel)
        Exit Function
    End If
    boundary = doc.Range(Start:=position, End:=position).Paragraphs(1).Range.Start
    If VTV2ImportCacheUsable(doc, position) Then
        If VT_WORD_V2_IMPORT_HEADINGS_LEVEL = targetLevel And _
           Not VT_WORD_V2_IMPORT_HEADINGS Is Nothing Then
            usable = (VT_WORD_V2_IMPORT_HEADINGS_END >= 0 And _
                VT_WORD_V2_IMPORT_HEADINGS_END <= boundary)
        End If
    End If
    If usable Then
        Set cached = VT_WORD_V2_IMPORT_HEADINGS
        For Each entry In VTV2CollectHeadingsBetween( _
            doc, targetLevel, VT_WORD_V2_IMPORT_HEADINGS_END, boundary)
            cached.Add entry
        Next entry
    Else
        Set cached = VTV2CollectHeadingsBetween( _
            doc, targetLevel, doc.Content.Start, boundary)
    End If
    Set VT_WORD_V2_IMPORT_HEADINGS = cached
    VT_WORD_V2_IMPORT_HEADINGS_END = boundary
    VT_WORD_V2_IMPORT_HEADINGS_LEVEL = targetLevel
    VTV2RememberImportPosition doc, position
    Set result = New Collection
    For Each entry In cached
        result.Add entry
    Next entry
    For Each entry In VTV2CollectHeadingsBetween( _
        doc, targetLevel, boundary, doc.Content.End)
        result.Add entry
    Next entry
    Set VTV2PlanHeadings = result
End Function

' --- numbering context for a new number (Windows PlanInsertion) --------------

' The document's numbering style: level 0 (continuous), 1 (chapter) or 2
' (section) and the separator "." or "-".
Public Sub VTV2DocumentFormat( _
    ByVal doc As Document, ByRef level As Long, ByRef separator As String)

    Dim numberingMode As String
    Dim separatorText As String

    VTResolveEquationNumbering doc, numberingMode, separatorText
    level = 0
    separator = ""
    If numberingMode = VT_WORD_NUMBERING_MODE_CHAPTER Then level = 1
    If numberingMode = VT_WORD_NUMBERING_MODE_SECTION Then level = 2
    If level > 0 Then
        separator = IIf(separatorText = ".", ".", "-")
    End If
End Sub

Public Function VTEquationStartsNumberScope( _
    ByVal target As Range, ByVal imageFormula As Boolean) As Boolean

    Dim documentObject As Document
    Dim shape As InlineShape
    Dim math As OMath
    Dim paragraph As Paragraph
    Dim formulaId As String
    Dim displayMode As String
    Dim numbered As Boolean
    Dim previousEnd As Long
    Dim targetLevel As Long
    Dim candidateLevel As Long
    Dim between As Range
    Dim exact As Range
    Dim roots As Collection
    Dim rootIndex As Long

    If target Is Nothing Then Exit Function
    Set documentObject = target.Document
    targetLevel = VTEquationNumberingRestartLevel(documentObject)
    If targetLevel < 1 Then Exit Function

    previousEnd = -1
    If imageFormula Then
        For Each shape In documentObject.InlineShapes
            If shape.Range.Start >= target.Start Then Exit For
            If VTTryParseFormulaReference( _
               shape.Title, formulaId, displayMode, numbered) Then
                If numbered And displayMode = "block" Then
                    previousEnd = shape.Range.Paragraphs(1).Range.End
                End If
            End If
        Next shape
    Else
        ' The last numbered equation before target decides; walk the
        ' SEQ-bearing equations backwards and stop at the first proven one.
        Set roots = VTNumberFieldNativeMathRoots( _
            documentObject, documentObject.Content.Start, target.Start)
        For rootIndex = roots.Count To 1 Step -1
            Set exact = roots(rootIndex)
            If exact.Start < target.Start Then
                If Not VTPureNativeNumberField(exact) Is Nothing Then
                    previousEnd = exact.Paragraphs(1).Range.End
                    Exit For
                End If
            End If
        Next rootIndex
    End If

    If previousEnd < 0 Then
        VTEquationStartsNumberScope = True
        Exit Function
    End If
    If previousEnd >= target.Start Then Exit Function
    Set between = documentObject.Range(previousEnd, target.Start)
    For Each paragraph In between.Paragraphs
        candidateLevel = paragraph.OutlineLevel
        If candidateLevel >= 1 And candidateLevel <= targetLevel Then
            VTEquationStartsNumberScope = True
            Exit Function
        End If
    Next paragraph
End Function

Public Function VTNativeHeadingLevelAtRange(ByVal target As Range) As Long
    If target Is Nothing Then Exit Function
    ' Chapter/section mode is a document contract, including before the first
    ' heading. The explicit zero scope (0 or 0.0) keeps the selected format
    ' visible instead of silently falling back to sequential numbering.
    VTNativeHeadingLevelAtRange = _
        VTEquationNumberingRestartLevel(target.Document)
End Function
