Attribute VB_Name = "sfc"
Option Explicit


Public Declare Function SfcGetNextProtectedFile Lib "sfc.dll" (ByVal RpcHandle As Long, ProtFileData As PROTECTED_FILE_DATA) As Long
'Public Declare Function SfcIsFileProtected Lib "sfc.dll" (ByVal RpcHandle As Long, ByVal ProtFileName As String) As Long

    Public Type PROTECTED_FILE_DATA
        FileName As String * MAX_PATH
        FileNumber As Long
    End Type
