Attribute VB_Name = "VTRibbonCallbacks"
Option Explicit

Public Sub VTWordRibbonInline(ByVal control As IRibbonControl)
    VisualTeX_CreateInline
End Sub

Public Sub VTWordRibbonDisplay(ByVal control As IRibbonControl)
    VisualTeX_CreateDisplay
End Sub

Public Sub VTWordRibbonEdit(ByVal control As IRibbonControl)
    VisualTeX_EditSelected
End Sub

Public Sub VTWordRibbonConvertNative(ByVal control As IRibbonControl)
    VisualTeX_ConvertSelectedToNativeEquation
End Sub

Public Sub VTWordRibbonNumbering(ByVal control As IRibbonControl)
    VisualTeX_UpdateEquationNumbers
End Sub

Public Sub VTWordRibbonOpen(ByVal control As IRibbonControl)
    VisualTeX_OpenApplication
End Sub
