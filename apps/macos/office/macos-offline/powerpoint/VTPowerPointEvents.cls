VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VTPowerPointEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As PowerPoint.Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WindowBeforeDoubleClick(ByVal Sel As PowerPoint.Selection, Cancel As Boolean)
    Dim selectedShape As PowerPoint.Shape

    On Error GoTo IgnoreDoubleClick
    If Sel Is Nothing Then Exit Sub
    If Sel.Type <> ppSelectionShapes Then Exit Sub
    If Sel.ShapeRange.Count <> 1 Then Exit Sub
    Set selectedShape = Sel.ShapeRange(1)
    If Not VTIsVisualTeXPowerPointShape(selectedShape) Then Exit Sub

    Cancel = True
    VisualTeX_EditShape selectedShape
    Exit Sub

IgnoreDoubleClick:
    ' Preserve PowerPoint's ordinary double-click behavior for every shape that
    ' is not a valid VisualTeX formula.
End Sub
