Attribute VB_Name = "VTErrorHandling"
Option Explicit

Public Sub VTShowError(ByVal operationName As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    Dim message As String
    message = "VisualTeX " & operationName & " failed."
    If Len(errorDescription) > 0 Then message = message & vbCrLf & vbCrLf & errorDescription
    If errorNumber <> 0 Then message = message & vbCrLf & vbCrLf & "Error " & CStr(errorNumber)
    MsgBox message, vbExclamation Or vbOKOnly, "VisualTeX"
End Sub

Public Sub VTShowInformation(ByVal message As String)
    MsgBox message, vbInformation Or vbOKOnly, "VisualTeX"
End Sub

Public Function VTBoundedIdentity(ByVal value As String) As String
    value = Replace$(value, vbCr, " ")
    value = Replace$(value, vbLf, " ")
    value = Replace$(value, Chr$(0), "")
    If Len(value) > 2048 Then value = Left$(value, 2048)
    VTBoundedIdentity = value
End Function
