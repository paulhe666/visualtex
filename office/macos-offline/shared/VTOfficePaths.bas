Attribute VB_Name = "VTOfficePaths"
Option Explicit

Private Const VT_OFFICE_GROUP_CONTAINER As String = "/Library/Group Containers/UBF8T346G9.Office/VisualTeX"

Public Function VTApplicationSupportRoot() As String
    Dim homePath As String
    Dim sandboxMarker As Long

    homePath = Environ$("HOME")
    sandboxMarker = InStr(1, homePath, "/Library/Containers/", vbTextCompare)
    If sandboxMarker > 1 Then
        homePath = Left$(homePath, sandboxMarker - 1)
    End If
    If Len(homePath) = 0 Or Left$(homePath, 1) <> "/" Then
        Err.Raise vbObjectError + 7100, "VisualTeX", "Unable to resolve the macOS home directory."
    End If

    ' Microsoft Office for Mac is sandboxed. Word and PowerPoint can both
    ' read and write their signed Office application group, while arbitrary
    ' user Application Support paths are not available to VBA.
    VTApplicationSupportRoot = homePath & VT_OFFICE_GROUP_CONTAINER
End Function
