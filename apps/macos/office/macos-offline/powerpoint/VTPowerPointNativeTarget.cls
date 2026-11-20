VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VTPowerPointNativeTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Session-only handles. None of these values is written into the presentation.
Public SessionId As String
Public Owner As PowerPoint.Presentation
Public OwnerWindow As PowerPoint.DocumentWindow
Public TargetSlide As PowerPoint.Slide
Public TargetShape As PowerPoint.Shape
Public SlideId As Long
Public ShapeId As Long
Public TextStart As Long
Public TextLength As Long
Public OriginalText As String
Public FontSizePt As Double
Public ReplaceStart As Long
Public ReplaceLength As Long
Public IsEdit As Boolean
Public Completed As Boolean
Public PrefixBreak As Boolean
Public SuffixBreak As Boolean
' An equation in a table cell: the table, and the cell whose text shape is TargetShape.
Public TableShape As PowerPoint.Shape
Public CellRow As Long
Public CellColumn As Long
