Attribute VB_Name = "VTErrorHandling"
Option Explicit

' An AI agent's operation (VisualTeX_AgentRun) is running: messages go to the agent, not into dialogs.
Public VT_AGENT_MODE As Boolean
Public VT_AGENT_ERRORS As String
Public VT_AGENT_NOTES As String
' The formula session an agent's command started: VTWriteAndLaunchSession hands its request to the app
' (VT_AGENT_REQUEST_JSON), which draws the agent's LaTeX in its hidden editor and applies it as Insert does; a
' document operation (redraw, restore, import) opens its window as usual and the app waits for it.
Public VT_AGENT_SESSION_ID As String
Public VT_AGENT_REQUEST_JSON As String

Public Sub VTShowError(ByVal operationName As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    Dim message As String
    message = "VisualTeX " & operationName & " failed."
    If Len(errorDescription) > 0 Then message = message & vbCrLf & vbCrLf & errorDescription
    If errorNumber <> 0 Then message = message & vbCrLf & vbCrLf & "Error " & CStr(errorNumber)
    If VT_AGENT_MODE Then
        If Len(VT_AGENT_ERRORS) > 0 Then VT_AGENT_ERRORS = VT_AGENT_ERRORS & vbLf
        VT_AGENT_ERRORS = VT_AGENT_ERRORS & Replace$(message, vbCrLf & vbCrLf, " ")
        Exit Sub
    End If
    MsgBox message, vbExclamation Or vbOKOnly, "VisualTeX"
End Sub

Public Sub VTShowInformation(ByVal message As String)
    If VT_AGENT_MODE Then
        If Len(VT_AGENT_NOTES) > 0 Then VT_AGENT_NOTES = VT_AGENT_NOTES & vbLf
        VT_AGENT_NOTES = VT_AGENT_NOTES & message
        Exit Sub
    End If
    MsgBox message, vbInformation Or vbOKOnly, "VisualTeX"
End Sub

' A callback the app runs (VisualTeX_ApplyPendingResult) reports its failure in the session folder, where the app
' reads it, instead of raising it: Office for Mac shows an error left by a macro run through AppleScript in a
' modal dialog, which stops the app and Office alike until someone closes it.
Public Function VTReportCallbackFailure( _
    ByVal sessionId As String, ByVal errorNumber As Long, ByVal errorDescription As String) As Boolean

    On Error GoTo Unreported
    If Len(sessionId) = 0 Then Exit Function
    VTWriteTextAtomic VTSessionDirectory(sessionId) & "/callback-error.txt", _
        CStr(errorNumber) & vbLf & errorDescription
    VTReportCallbackFailure = True
Unreported:
End Function

Public Function VTBoundedIdentity(ByVal value As String) As String
    value = Replace$(value, vbCr, " ")
    value = Replace$(value, vbLf, " ")
    value = Replace$(value, Chr$(0), "")
    If Len(value) > 2048 Then value = Left$(value, 2048)
    VTBoundedIdentity = value
End Function
