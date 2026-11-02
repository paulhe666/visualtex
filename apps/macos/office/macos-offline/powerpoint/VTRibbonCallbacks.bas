Attribute VB_Name = "VTRibbonCallbacks"
Option Explicit

Public Sub VTPowerPointRibbonInlineOmml(ByVal control As IRibbonControl)
    VisualTeX_NewInlineNativeEquation
End Sub

Public Sub VTPowerPointRibbonDisplayOmml(ByVal control As IRibbonControl)
    VisualTeX_NewDisplayNativeEquation
End Sub

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

Public Sub VTPowerPointRibbonRedrawOmmlSelection(ByVal control As IRibbonControl)
    VisualTeX_RedrawSelectionToOmml
End Sub

Public Sub VTPowerPointRibbonRedrawOmmlSlide(ByVal control As IRibbonControl)
    VisualTeX_RedrawSlideToOmml
End Sub

Public Sub VTPowerPointRibbonRedrawOmmlAll(ByVal control As IRibbonControl)
    VisualTeX_RedrawPresentationToOmml
End Sub

Public Sub VTPowerPointRibbonRedrawImageSelection(ByVal control As IRibbonControl)
    VisualTeX_RedrawSelectionToImage
End Sub

Public Sub VTPowerPointRibbonRedrawImageSlide(ByVal control As IRibbonControl)
    VisualTeX_RedrawSlideToImage
End Sub

Public Sub VTPowerPointRibbonRedrawImageAll(ByVal control As IRibbonControl)
    VisualTeX_RedrawPresentationToImage
End Sub

Public Sub VTPowerPointRibbonToLatexSelection(ByVal control As IRibbonControl)
    VisualTeX_RestoreSelectionToLatex
End Sub

Public Sub VTPowerPointRibbonToLatexSlide(ByVal control As IRibbonControl)
    VisualTeX_RestoreSlideToLatex
End Sub

Public Sub VTPowerPointRibbonToLatexAll(ByVal control As IRibbonControl)
    VisualTeX_RestorePresentationToLatex
End Sub
