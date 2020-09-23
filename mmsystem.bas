Attribute VB_Name = "mmsystem"
Option Explicit

Public Declare Function auxGetDevCaps Lib "winmm.dll" Alias "auxGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As AUXCAPS, ByVal uSize As Long) As Long
Public Declare Function auxGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long


    Public Const MAXPNAMELEN = 32  '  max product name length (including NULL)

    
    Public Type AUXCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        wTechnology As Integer
        dwSupport As Long
    End Type


Public Function mciError(lngError)
    Dim errDescription As String
    Dim lenError As Long
    
    errDescription = Space$(128)
    
    apiError = mciGetErrorString(lngError, errDescription, Len(errDescription))
    
    If errDescription <> "" Then
        If errMsg = True Then
            MsgBox errDescription, vbExclamation, "Error"
        End If
    End If
End Function
