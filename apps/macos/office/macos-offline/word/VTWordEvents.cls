VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VTWordEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WindowSelectionChange(ByVal Sel As Word.Selection)
    On Error Resume Next
    If Sel Is Nothing Or VTWordInternalMutationActive() Then Exit Sub
    ' Word inserts Return before the external SEQ helper when the caret is placed
    ' at the visible end of a numbered image row. Repair only that exact empty-line
    ' mutation first; the helper remains adjacent to its formula and the caret moves
    ' to the equivalent ordinary paragraph after the helper.
    If VTRepairNumberedImageReturnAfterFormula(Sel.Range) Then
        VTInvalidateWordImageFontSizeControl
        Exit Sub
    End If
    ' Word for Mac can create a private _Ref at the framed caption anchor rather
    ' than around the helper's pure number. Repair only private REF fields in the
    ' currently selected body paragraph, then leave F9 and hyperlink semantics
    ' entirely native.
    If VTRepairSelectedWordNativeEquationCrossReference(Sel.Range) Then
        VTInvalidateWordImageFontSizeControl
        Exit Sub
    End If
    VTCleanupEmptyVisualTeXImageMacroButtons Sel.Range
    VisualTeX_StabilizeImageEquationNumberSelection Sel.Range
    VisualTeX_SynchronizeSelectedImageFormulaSize Sel
    VTInvalidateWordImageFontSizeControl
End Sub

Private Sub App_DocumentOpen(ByVal Doc As Word.Document)
    On Error Resume Next
    VTEnsureImageMacroMigrationScheduled
    VTInvalidateWordEquationNumberingFormatControl
    On Error GoTo 0
End Sub

Private Sub App_NewDocument(ByVal Doc As Word.Document)
    On Error Resume Next
    VTEnsureImageMacroMigrationScheduled
    VTInvalidateWordEquationNumberingFormatControl
    On Error GoTo 0
End Sub

Private Sub App_WindowActivate( _
    ByVal Doc As Word.Document, _
    ByVal Wn As Word.Window)

    On Error Resume Next
    ' Word for Mac can re-quantize an SVG InlineShape while reopening a document
    ' or reactivating its window without raising WindowSelectionChange. Run the
    ' same selected-formula synchronization here so the mathematical baseline
    ' is repaired from persisted geometry before the next edit. The resolver is
    ' local to the current selection and returns immediately for ordinary text.
    If Not VTWordInternalMutationActive() Then _
        VisualTeX_SynchronizeSelectedImageFormulaSize Selection
    VTInvalidateWordEquationNumberingFormatControl
    On Error GoTo 0
End Sub

Private Sub App_WindowBeforeDoubleClick(ByVal Sel As Word.Selection, Cancel As Boolean)
    VTBeginWordDoubleClickTraceContext
    On Error GoTo IgnoreDoubleClick
    VTTraceWordDoubleClick "window-before-double-click-enter", Sel, ""
    If Sel Is Nothing Then
        VTTraceWordDoubleClick _
            "window-before-double-click-skip", Sel, _
            "reason=no-selection"
        GoTo DoubleClickFinished
    End If
    Cancel = VTHandleWordBeforeDoubleClick(Sel)
    VTTraceWordDoubleClick _
        "window-before-double-click-result", Sel, _
        "cancel=" & CStr(Cancel)
    GoTo DoubleClickFinished

IgnoreDoubleClick:
    VTTraceWordDoubleClick _
        "window-before-double-click-error", Sel, _
        "error=" & CStr(Err.Number) & ":" & Err.Description
    ' Ordinary Word double-click behavior must remain untouched when the
    ' clicked object is not a valid VisualTeX formula.
DoubleClickFinished:
    VTEndWordDoubleClickTraceContext
End Sub
