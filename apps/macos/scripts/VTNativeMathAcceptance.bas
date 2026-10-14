Attribute VB_Name = "VTNativeMathAcceptance"
Option Explicit

' Imported only into the owned acceptance PPTM; excluded from the shipped PPAM.
Public Sub VTNativeAcceptancePrepareCase(ByVal caseName As String, ByVal sessionId As String)
    Dim p As Presentation
    Dim s As Slide
    Dim box As Shape
    Dim selectedText As TextRange
    Dim target As VTPowerPointNativeTarget
    Dim text As String
    Dim start As Long
    Dim length As Long
    Dim size As Single

    Set p = Presentations("NativeOmmlAcceptance.pptm")
    If InStr(p.FullName, "/visualtexmac/apps/macos/build-logs/powerpoint-native-omml/") = 0 Then Err.Raise 5, , "Not the owned acceptance presentation."
    p.Windows(1).Activate
    Set s = p.Slides.Add(p.Slides.Count + 1, ppLayoutBlank)
    p.Windows(1).View.GotoSlide s.SlideIndex
    If caseName = "standalone_inline" Or caseName = "standalone_display" Then
        ActiveWindow.Selection.Unselect
        Set target = VTNativeCaptureTarget(sessionId)
        If Not target.TargetShape Is Nothing Then Err.Raise 5, , "Expected a standalone target."
        Exit Sub
    End If
    text = "before X after"
    start = 8
    length = 1
    size = 24!
    Select Case caseName
        Case "inline_caret": text = "before after": length = 0
        Case "inline_start": text = "after": start = 1: length = 0
        Case "inline_end": text = "before": start = 7: length = 0
        Case "inline_empty", "display_empty": text = "": start = 1: length = 0
        Case "display_existing": text = "before" & vbCr & "X" & vbCr & "after"
        Case "display_start": text = "X after": start = 1
        Case "display_end": text = "before X"
        Case "unicode": text = ChrW(&H524D) & ChrW(&H6587) & " X " & ChrW(&H540E) & ChrW(&H6587): start = 4
        Case "multiline_replace": text = "before ONE" & vbCr & "TWO after": length = 7
        Case "font_size": size = 36!
    End Select
    Set box = s.Shapes.AddTextbox(msoTextOrientationHorizontal, 60!, 120!, 600!, 210!)
    box.Name = "Case_" & caseName
    box.TextFrame.TextRange.Text = text
    box.TextFrame.TextRange.Font.Size = size
    If caseName = "format_preservation" Or caseName = "display_bullets" Then
        box.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignRight
        box.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 11!
        box.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 13!
        box.TextFrame.TextRange.Characters(1, 7).Font.Bold = msoTrue
        box.TextFrame.TextRange.Characters(10, 5).Font.Italic = msoTrue
        If caseName = "display_bullets" Then box.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = msoTrue
    End If
    Set selectedText = box.TextFrame.TextRange.Characters(start, length)
    selectedText.Select
    Set target = VTNativeCaptureTarget(sessionId)
    If target.TargetShape Is Nothing Then Err.Raise 5, , "Text selection became standalone."
    If target.TextStart <> start Or target.TextLength <> length Then
        Err.Raise 5, , "Unexpected PowerPoint caret: " & CStr(target.TextStart) & ":" & CStr(target.TextLength)
    End If
    If target.FontSizePt <> size Then Err.Raise 5, , "Native math failed to inherit the selected font size."
    If caseName = "stale" Then box.TextFrame.TextRange.InsertAfter " changed"
End Sub

Public Sub VTNativeAcceptanceReport(ByVal sessionId As String)
    Dim p As Presentation
    Dim s As Slide
    Dim box As Shape
    Dim summary As String
    Set p = Presentations("NativeOmmlAcceptance.pptm")
    If InStr(p.FullName, "/visualtexmac/apps/macos/build-logs/powerpoint-native-omml/") = 0 Then Err.Raise 5
    Set s = p.Slides(p.Slides.Count)
    summary = CStr(s.SlideIndex) & "|" & CStr(s.Shapes.Count) & "|"
    For Each box In s.Shapes
        If box.HasTextFrame = msoTrue Then summary = summary & box.TextFrame.TextRange.Text
    Next box
    VTWriteTextAtomic VTSessionDirectory(sessionId) & "/acceptance-report.txt", summary
End Sub

Public Sub VTNativeEditPrepare(ByVal caseName As String, ByVal sessionId As String)
    Dim p As Presentation
    Dim copySlides As SlideRange
    Dim s As Slide
    Dim box As Shape
    Dim zone As Object
    Dim resolved As Object
    Dim resolvedShape As Shape
    Dim target As VTPowerPointNativeTarget
    Dim fullText As Object
    Dim index As Long
    Dim report As String
    Dim usePoint As Boolean
    Dim x As Double, y As Double
    Dim windowObject As Object
    On Error GoTo Failed
    Set p = Presentations("NativeOmmlAcceptance.pptm")
    If InStr(p.FullName, "/visualtexmac/apps/macos/build-logs/powerpoint-native-omml/") = 0 Then Err.Raise 5, , "Not the owned test project."
    Set copySlides = p.Slides(IIf(caseName = "display", 16, 79)).Duplicate
    Set s = copySlides(1)
    s.MoveTo p.Slides.Count
    p.Windows(1).Activate
    p.Windows(1).View.GotoSlide s.SlideIndex
    Set windowObject = p.Windows(1)
    Set box = s.Shapes(1)
    Set fullText = box.TextFrame2.TextRange
    If Left$(caseName, 7) = "corpus_" Or caseName = "merged_inside" Then
        box.TextFrame.TextRange.Text = "before X after"
        box.TextFrame.TextRange.Characters(8, 1).Paste
        Set fullText = box.TextFrame2.TextRange
        If caseName = "merged_inside" Then
            Set zone = fullText.MathZones(1, 1)
            box.TextFrame.TextRange.Characters(zone.Start + 1, 0).Paste
            Set fullText = box.TextFrame2.TextRange
        End If
    End If
    If caseName = "multiple" Or caseName = "ambiguous" Or caseName = "merged" Then
        box.TextFrame.TextRange.Text = "before X between Y after"
        box.TextFrame.TextRange.Characters(8, 1).Paste
        index = InStr(box.TextFrame.TextRange.Text, "Y")
        box.TextFrame.TextRange.Characters(index, 1).Paste
        Set fullText = box.TextFrame2.TextRange
        If caseName = "merged" Then
            Set zone = fullText.MathZones(1, 1)
            index = zone.Start + zone.Length
            Dim second As Object
            Set second = fullText.MathZones(2, 1)
            box.TextFrame.TextRange.Characters(index, second.Start - index).Delete
            Set fullText = box.TextFrame2.TextRange
        End If
    End If
    report = "slide=" & CStr(s.SlideIndex) & vbLf & "zones=" & CStr(fullText.MathZones.Count) & vbLf
    Set zone = fullText.MathZones(IIf(caseName = "multiple", 2, 1), 1)
    If caseName = "shape" Or caseName = "ambiguous" Then
        box.Select
    ElseIf caseName = "ordinary" Then
        box.TextFrame.TextRange.Characters(1, 3).Select
    ElseIf caseName = "partial" Then
        box.TextFrame.TextRange.Characters(zone.Start + 1, 2).Select
    Else
        box.TextFrame.TextRange.Characters(zone.Start + 1, 0).Select
    End If
    If caseName = "point" Or caseName = "miss" Then
        usePoint = True
        x = windowObject.PointsToScreenPixelsX(zone.BoundLeft + zone.BoundWidth / 2#)
        y = windowObject.PointsToScreenPixelsY(zone.BoundTop + zone.BoundHeight / 2#)
        If caseName = "miss" Then x = windowObject.PointsToScreenPixelsX(0)
    End If
    If VTTryResolveNativeMathZone(ActiveWindow.Selection, resolvedShape, resolved, usePoint, x, y) Then
        report = report & "resolved=" & CStr(resolved.Start) & ":" & CStr(resolved.Length) & vbLf
        Set target = VTNativeCaptureEditTarget(sessionId, resolvedShape, resolved)
        VTWriteRequest sessionId, VTNativeEditRequest(target)
    Else
        report = report & "resolved=none" & vbLf
    End If
Finished:
    VTWriteTextAtomic VTSessionDirectory(sessionId) & "/native-edit-prepare.txt", report
    Exit Sub
Failed:
    report = report & "error=" & CStr(Err.Number) & ":" & Err.Description
    Resume Finished
End Sub

Public Sub VTNativeUiUnloadCandidate()
    Dim item As AddIn
    For Each item In Application.AddIns
        If LCase$(item.Name) = "visualtexnativeeditvalidation" Then item.Loaded = msoFalse
    Next item
End Sub

Public Sub VTNativeUiLoadCandidate()
    Dim item As AddIn
    Dim candidate As AddIn
    Dim root As String
    Dim current As Presentation
    Set current = Presentations("NativeOmmlAcceptance.pptm")
    root = Left$(current.FullName, InStr(current.FullName, "/build-logs/") - 1)
    If Right$(root, 12) <> "/apps/macos" And InStr(root, "/visualtexmac/apps/macos") = 0 Then Err.Raise 5
    For Each item In Application.AddIns
        If LCase$(item.Name) = "visualtex" Or LCase$(item.Name) = "visualtex.ppam" Then item.Loaded = msoFalse
    Next item
    Set candidate = Application.AddIns.Add(root & "/build-logs/powerpoint-native-omml/VisualTeXNativeEditValidation.ppam")
    candidate.Loaded = msoTrue
End Sub

Public Sub VTNativeReplaceBodyProbe(ByVal mode As String)
    Dim p As Presentation
    Dim s As Slide
    Dim box As Shape
    Dim zone As Object, body As Object, pasted As Object
    Dim report As String
    Set p = Presentations("NativeOmmlAcceptance.pptm")
    p.Windows(1).Activate
    Set s = p.Slides.Add(p.Slides.Count + 1, ppLayoutBlank)
    p.Windows(1).View.GotoSlide s.SlideIndex
    Set box = s.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 100, 600, 120)
    box.TextFrame.TextRange.Text = "before X after"
    box.TextFrame.TextRange.Characters(8, 1).Paste
    Set zone = box.TextFrame2.TextRange.MathZones(1, 1)
    report = "before=" & box.TextFrame.TextRange.Text & vbLf & "zone=" & zone.Start & ":" & zone.Length & vbLf
    If mode = "context" Then
        Set body = box.TextFrame2.TextRange.Characters(zone.Start - 1, zone.Length + 2)
        body.Copy
    ElseIf mode = "delete" Then
        Dim start As Long
        start = zone.Start
        zone.Delete
        report = report & "afterDelete=" & box.TextFrame.TextRange.Text & vbLf
        Set body = box.TextFrame2.TextRange.Characters(start, 0)
    ElseIf mode = "inner" Then
        Set body = zone.Characters(2, zone.Length - 2)
    Else
        Set body = zone
    End If
    report = report & "target=" & body.Start & ":" & body.Length & ":" & body.Text & vbLf
    Set pasted = body.Paste
    report = report & "after=" & box.TextFrame.TextRange.Text & vbLf & "slide=" & s.SlideIndex & vbLf
    VTWriteTextAtomic VTApplicationSupportRoot() & "/Tests/native-body-" & mode & ".txt", report
    p.Save
End Sub

Public Sub VTNativeIndexProbe()
    Dim p As Presentation
    Dim box As Shape
    Dim first As Object, second As Object, zone As Object
    Dim selected As Object
    Dim report As String
    Dim index As Long
    Set p = Presentations("NativeOmmlAcceptance.pptm")
    p.Windows(1).Activate
    p.Windows(1).View.GotoSlide 7
    Set box = p.Slides(7).Shapes(1)
    Set first = box.TextFrame.TextRange
    Set second = box.TextFrame2.TextRange
    Set zone = second.MathZones(1, 1)
    report = "TEXT1=" & CStr(first.Length) & ":" & first.Text & vbLf & "TEXT2=" & CStr(second.Length) & ":" & second.Text & vbLf
    report = report & "ZONE=" & CStr(zone.Start) & ":" & CStr(zone.Length) & ":" & zone.Text & vbLf
    report = report & "SLICE1=" & first.Characters(zone.Start, zone.Length).Text & vbLf
    report = report & "SLICE2=" & second.Characters(zone.Start, zone.Length).Text & vbLf
    zone.Select
    Set selected = ActiveWindow.Selection
    report = report & "SEL1=" & CStr(selected.TextRange.Start) & ":" & CStr(selected.TextRange.Length) & ":" & selected.TextRange.Text & vbLf
    On Error Resume Next
    report = report & "SEL2=" & CStr(selected.TextRange2.Start) & ":" & CStr(selected.TextRange2.Length) & ":" & selected.TextRange2.Text & vbLf
    report = report & "SEL2ERR=" & CStr(Err.Number) & ":" & Err.Description & vbLf
    For index = 1 To first.Length
        report = report & "CHAR" & CStr(index) & "=" & first.Characters(index, 1).Text & "|" & second.Characters(index, 1).Text & vbLf
    Next index
    VTWriteTextAtomic VTApplicationSupportRoot() & "/Tests/native-index-probe.txt", report
End Sub

Public Sub VTNativeUiReport()
    Dim item As AddIn
    Dim text As String
    Dim selected As Selection
    Dim target As Shape
    Dim zone As Object
    For Each item In Application.AddIns
        text = text & item.Name & "|" & item.FullName & "|" & CStr(item.Loaded) & vbLf
    Next item
    Set selected = ActiveWindow.Selection
    text = text & "selectionType=" & CStr(selected.Type) & vbLf
    On Error Resume Next
    text = text & "range=" & CStr(selected.TextRange.Start) & ":" & CStr(selected.TextRange.Length) & vbLf
    text = text & "resolve=" & CStr(VTTryResolveNativeMathZone(selected, target, zone)) & vbLf
    If Not zone Is Nothing Then text = text & "math=" & CStr(zone.Start) & ":" & CStr(zone.Length) & vbLf
    VTWriteTextAtomic VTApplicationSupportRoot() & "/Tests/native-ui-report.txt", text
End Sub

Public Sub VTNativeUiSelect(ByVal slideIndexText As String)
    Dim p As Presentation
    Dim s As Slide
    Dim zone As Object
    Dim win As Object
    Dim report As String
    Set p = Presentations("NativeOmmlAcceptance.pptm")
    p.Windows(1).Activate
    Set s = p.Slides(CLng(slideIndexText))
    p.Windows(1).View.GotoSlide s.SlideIndex
    Set zone = s.Shapes(1).TextFrame2.TextRange.MathZones(1, 1)
    s.Shapes(1).TextFrame.TextRange.Characters(zone.Start + 1, 0).Select
    Set win = p.Windows(1)
    report = CStr(win.PointsToScreenPixelsX(zone.BoundLeft + zone.BoundWidth / 2#)) & "|" & _
             CStr(win.PointsToScreenPixelsY(zone.BoundTop + zone.BoundHeight / 2#))
    VTWriteTextAtomic VTApplicationSupportRoot() & "/Tests/native-ui-point.txt", report
End Sub

Public Sub VTNativeEditMutateStyle()
    Dim p As Presentation
    Set p = Presentations("NativeOmmlAcceptance.pptm")
    p.Slides(p.Slides.Count).Shapes(1).TextFrame2.TextRange.MathZones(1, 1).Font.Bold = msoTrue
End Sub

Public Sub VTNativeEditProbe()
    Dim p As Presentation
    Dim s As Slide
    Dim box As Shape
    Dim fullText As Object
    Dim zones As Object
    Dim zone As Object
    Dim part As Object
    Dim i As Long
    Dim report As String
    Dim n As Long
    Dim errorText As String
    On Error GoTo Failed
    Set p = Presentations("NativeOmmlAcceptance.pptm")
    p.Windows(1).Activate
    Set s = p.Slides(7)
    p.Windows(1).View.GotoSlide s.SlideIndex
    Set box = s.Shapes(1)
    Set fullText = box.TextFrame2.TextRange
    report = "full=" & CStr(fullText.Start) & ":" & CStr(fullText.Length) & ":" & fullText.Text & vbLf
    Set zones = fullText.MathZones
    report = report & "zones=" & CStr(zones.Count) & ":" & CStr(zones.Start) & ":" & CStr(zones.Length) & ":" & zones.Text & vbLf
    For i = 1 To 4
        On Error Resume Next
        Set zone = Nothing
        Err.Clear
        Set zone = fullText.MathZones(i, 1)
        If Err.Number = 0 And Not zone Is Nothing Then
            report = report & "zone" & CStr(i) & "=" & CStr(zone.Start) & ":" & CStr(zone.Length) & ":" & zone.Text & vbLf
        Else
            report = report & "zone" & CStr(i) & "error=" & CStr(Err.Number) & ":" & Err.Description & vbLf
        End If
        On Error GoTo Failed
    Next i
    For i = 1 To fullText.Length + 1
        On Error Resume Next
        Set part = fullText.Characters(i, 0)
        Set zone = Nothing
        Err.Clear
        Set zone = part.MathZones
        If Err.Number = 0 And Not zone Is Nothing Then
            report = report & "caret" & CStr(i) & "=" & CStr(zone.Start) & ":" & CStr(zone.Length) & ":" & zone.Text & vbLf
        Else
            report = report & "caret" & CStr(i) & "error=" & CStr(Err.Number) & ":" & Err.Description & vbLf
        End If
        On Error GoTo Failed
    Next i
    Set zones = p.Slides(25).Shapes(1).TextFrame2.TextRange.MathZones
    report = report & "plainZones=" & CStr(zones.Count) & ":" & CStr(zones.Length) & vbLf
    On Error Resume Next
    Dim win As Object
    Set win = p.Windows(1)
    Err.Clear
    report = report & "screenOrigin=" & CStr(win.PointsToScreenPixelsX(0)) & ":" & CStr(win.PointsToScreenPixelsY(0)) & vbLf
    report = report & "screenError=" & CStr(Err.Number) & ":" & Err.Description & vbLf
    Set zone = fullText.MathZones(1, 1)
    report = report & "bounds=" & CStr(zone.BoundLeft) & ":" & CStr(zone.BoundTop) & ":" & CStr(zone.BoundWidth) & ":" & CStr(zone.BoundHeight) & vbLf
    On Error GoTo Failed
Finished:
    VTWriteTextAtomic VTApplicationSupportRoot() & "/Tests/native-edit-probe.txt", report
    Exit Sub
Failed:
    report = report & "FAILED=" & CStr(Err.Number) & ":" & Err.Description & vbLf
    Resume Finished
End Sub

Public Sub VTNativeAcceptanceApply()
    ' Resolve the callback in THIS compiled project, not the installed PPAM.
    VisualTeX_ApplyPendingResult
End Sub

Public Sub VTNativeAcceptanceSave()
    Dim p As Presentation
    Set p = Presentations("NativeOmmlAcceptance.pptm")
    If InStr(p.FullName, "/visualtexmac/apps/macos/build-logs/powerpoint-native-omml/") = 0 Then Err.Raise 5
    p.Save
End Sub
