Attribute VB_Name = "VTRibbonCallbacks"
Option Explicit

Public Sub VTPowerPointRibbonNew(ByVal control As IRibbonControl)
    VisualTeX_NewFormula
End Sub

Public Sub VTPowerPointRibbonEdit(ByVal control As IRibbonControl)
    VisualTeX_EditSelected
End Sub

Public Sub VTPowerPointRibbonDelete(ByVal control As IRibbonControl)
    VisualTeX_DeleteSelected
End Sub

Public Sub VTPowerPointRibbonOpen(ByVal control As IRibbonControl)
    VisualTeX_OpenApplication
End Sub
