Attribute VB_Name = "errors"
Option Explicit

Public Sub Errors(lngError As Long, apiFunction As String, Optional errDescription As String, Optional NoMsgBox As Boolean)
    errDescription = Space$(2048)
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, 0&, lngError, 0, errDescription, 2048, 0&
    
    errDescription = Trim$(errDescription)
    If errDescription = "" Then errDescription = "No description available."
    
    If NoMsgBox = False Then
        If errMsg = True Then
            MsgBox apiFunction & vbCrLf & vbCrLf & errDescription, vbExclamation, "Error"
        End If
    End If
End Sub

Public Sub Failed(strAPI As String)
    If errMsg = True Then
        If Err.LastDllError = 0 Then
            MsgBox strAPI & vbCrLf & vbCrLf & "Failed", vbExclamation, "Error"
        Else
            Errors Err.LastDllError, strAPI
        End If
    End If
End Sub
