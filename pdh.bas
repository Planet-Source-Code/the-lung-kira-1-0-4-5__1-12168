Attribute VB_Name = "pdh"
Option Explicit

Public Declare Function PdhCloseQuery Lib "PDH.DLL" (ByVal QueryHandle As Long) As Long
Public Declare Function PdhCollectQueryData Lib "PDH.DLL" (ByVal QueryHandle As Long) As Long
Public Declare Function PdhEnumObjects Lib "PDH.DLL" Alias "PdhEnumObjectsA" (ByVal szDataSource As String, ByVal szMachineName As String, ByVal mszObjectList As String, ByRef pcchBufferSize As Long, ByVal dwDetailLevel As Long, ByVal bRefresh As Boolean) As Long
Public Declare Function PdhEnumObjectItems Lib "PDH.DLL" Alias "PdhEnumObjectItemsA" (ByVal szDataSource As String, ByVal szMachineName As String, ByVal szObjectName As String, ByVal mszCounterList As String, ByVal pcchCounterListLength As Long, ByVal mszInstanceList As String, ByVal pcchInstanceListLength As Long, ByVal dwDetailLevel As Long, ByVal dwFlags As Long) As Long
Public Declare Function PdhOpenQuery Lib "PDH.DLL" (ByVal Reserved As Long, ByVal dwUserData As Long, ByRef hQuery As Long) As Long
Public Declare Function PdhRemoveCounter Lib "PDH.DLL" (ByVal CounterHandle As Long) As Long

Public Declare Function PdhVbAddCounter Lib "PDH.DLL" (ByVal QueryHandle As Long, ByVal CounterPath As String, ByRef CounterHandle As Long) As Long
Public Declare Function PdhVbGetOneCounterPath Lib "PDH.DLL" (ByVal PathString As String, ByVal PathLength As Long, ByVal DetailLevel As Long, ByVal CaptionString As String) As Long
Public Declare Function PdhVbGetDoubleCounterValue Lib "PDH.DLL" (ByVal CounterHandle As Long, ByRef CounterStatus As Long) As Double
Public Declare Function PdhVbOpenQuery Lib "PDH.DLL" (ByRef QueryHandle As Long) As Long
