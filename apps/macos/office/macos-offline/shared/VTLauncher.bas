Attribute VB_Name = "VTLauncher"
Option Explicit

Public Sub VTLaunchSession(ByVal hostName As String, ByVal sessionId As String)
    Dim scriptName As String
    Dim response As String

    If Not VTIsCanonicalUuid(sessionId) Then
        Err.Raise vbObjectError + 7300, "VisualTeX", "Invalid VisualTeX Session id."
    End If

    Select Case LCase$(hostName)
        Case "word": scriptName = "VisualTeXWord.scpt"
        Case "powerpoint": scriptName = "VisualTeXPowerPoint.scpt"
        Case Else
            Err.Raise vbObjectError + 7301, "VisualTeX", "Invalid VisualTeX Office host."
    End Select

#If Mac Then
    response = AppleScriptTask(scriptName, "OpenVisualTeXSession", sessionId)
#Else
    Err.Raise vbObjectError + 7302, "VisualTeX", "The VisualTeX offline add-in is available only on macOS."
#End If

    If Left$(response, 3) <> "ok|" Then
        Err.Raise vbObjectError + 7303, "VisualTeX", VTAppleScriptErrorMessage(response)
    End If
End Sub

Public Sub VTOpenApplication(ByVal hostName As String)
    Dim scriptName As String
    Dim response As String

    Select Case LCase$(hostName)
        Case "word": scriptName = "VisualTeXWord.scpt"
        Case "powerpoint": scriptName = "VisualTeXPowerPoint.scpt"
        Case Else
            Err.Raise vbObjectError + 7304, "VisualTeX", "Invalid VisualTeX Office host."
    End Select

#If Mac Then
    response = AppleScriptTask(scriptName, "OpenVisualTeXApplication", "")
#Else
    Err.Raise vbObjectError + 7305, "VisualTeX", "The VisualTeX offline add-in is available only on macOS."
#End If

    If Left$(response, 3) <> "ok|" Then
        Err.Raise vbObjectError + 7306, "VisualTeX", VTAppleScriptErrorMessage(response)
    End If
End Sub

Private Function VTAppleScriptErrorMessage(ByVal response As String) As String
    Dim fields() As String
    If Left$(response, 6) = "error|" Then
        fields = Split(response, "|")
        If UBound(fields) >= 2 Then
            VTAppleScriptErrorMessage = fields(2)
            Exit Function
        End If
    End If
    VTAppleScriptErrorMessage = "VisualTeX could not be opened."
End Function
