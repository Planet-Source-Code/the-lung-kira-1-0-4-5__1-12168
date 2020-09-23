Attribute VB_Name = "winreg"
Option Explicit


Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long


    Public Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_USER = &H80000001
    Public Const HKEY_LOCAL_MACHINE = &H80000002
    Public Const HKEY_USERS = &H80000003
    Public Const HKEY_PERFORMANCE_DATA = &H80000004
    Public Const HKEY_CURRENT_CONFIG = &H80000005
    Public Const HKEY_DYN_DATA = &H80000006
    

Public Function CreateKey(hKey As Long, strPath As String, Optional retDisposition As Long)
    Dim retKey As Long
    
    Dim SECURITY_ATTRIBUTES As SECURITY_ATTRIBUTES
    
    apiError = RegCreateKeyEx(hKey, strPath & Chr$(0), 0, Chr$(0), REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SECURITY_ATTRIBUTES, retKey, retDisposition) 'Creates key , if exists opens key
    If Not apiError = 0 Then
        Errors.Errors apiError, "RegCreateKeyEx"
        Exit Function
    End If
    
    apiError = RegCloseKey(retKey) 'Closes the current open key
    If Not apiError = 0 Then Errors.Errors apiError, "RegCloseKey"
End Function

Public Function DeleteKey(hKey As Long, strPath As String)
    apiError = RegDeleteKey(hKey, strPath)
    If Not apiError = 0 Then Errors.Errors apiError, "RegDeleteKey"
End Function

Public Function DeleteValue(hKey As Long, strPath As String, strValueName As String)
    Dim hCurKey As Long
    
    apiError = RegOpenKeyEx(hKey, strPath, vbNull, KEY_SET_VALUE, hCurKey)
    If Not apiError = 0 Then
        Errors.Errors apiError, "RegOpenKeyEx"
        Exit Function
    End If
    
    apiError = RegDeleteValue(hCurKey, strValueName)
    If Not apiError = 0 Then Errors.Errors apiError, "RegDeleteValue"
    
    apiError = RegCloseKey(hCurKey) 'Closes the current open key
    If Not apiError = 0 Then Errors.Errors apiError, "RegCloseKey"
End Function

Public Function EnumKeyEx(hKey As Long, strPath As String, strKeyName() As String, lngCount As Long)
    Dim hCurKey As Long
    Dim lenKeyName() As Long
    Dim FILETIME As FILETIME
    
    Do 'Cycle until all keys are read unless error occured
        RegOpenKeyEx hKey, strPath, 0, KEY_ENUMERATE_SUB_KEYS, hCurKey
        
        'Resizes array without destroying
        ReDim Preserve strKeyName(lngCount)
        ReDim Preserve lenKeyName(lngCount)
        
        strKeyName(lngCount) = Space$(1024) & Chr$(0) 'Null term
        lenKeyName(lngCount) = Len(strKeyName(lngCount))
        
        apiError = RegEnumKeyEx(hCurKey, lngCount, strKeyName(lngCount), lenKeyName(lngCount), 0&, 0&, 0&, FILETIME)
        
        'Send back out with out null term & padding
        If lenKeyName(lngCount) > 0 Then
            strKeyName(lngCount) = Left$(strKeyName(lngCount), lenKeyName(lngCount))
        End If
        
        lngCount = lngCount + 1 'Increment
        RegCloseKey hCurKey 'Closes the current open key
        
        If apiError > 0 Then
            Exit Do
        End If
    Loop While apiError <> ERROR_NO_MORE_ITEMS
    
    If Not apiError = ERROR_NO_MORE_ITEMS Then
        Errors.Errors apiError, "RegEnumKeyEx"
    End If
End Function

Public Function EnumValue(hKey As Long, strPath As String, strValueName() As String, lngCount As Long, lngValueType As Long)
    Dim hCurKey As Long
    Dim apiError As Long
    Dim lenValueName() As Long
    Dim lenMaxValueName As Long
    Dim FILETIME As FILETIME
    
    'Dim byDataBuffer(1024) As Byte
    Dim lenData As Long
    lenData = 4096
    
    Do 'Cycle until all values are read unless error occured
        RegOpenKeyEx hKey, strPath, 0&, KEY_READ, hCurKey
        RegQueryInfoKey hCurKey, 0&, 0&, 0&, 0&, 0&, 0&, 0&, lenMaxValueName, 0&, 0&, FILETIME

        'Resizes array without destroying
        ReDim Preserve strValueName(lngCount)
        ReDim Preserve lenValueName(lngCount)

        strValueName(lngCount) = Space$(lenMaxValueName) & Chr$(0) 'Null term
        lenValueName(lngCount) = Len(strValueName(lngCount))

        apiError = RegEnumValue(hCurKey, lngCount, strValueName(lngCount), lenValueName(lngCount), 0&, lngValueType, 0&, lenData)

       'Send back out with out null term & padding
        If lenValueName(lngCount) > 0 Then
            strValueName(lngCount) = Fix_NullTermStr(strValueName(lngCount))
        End If

        lngCount = lngCount + 1 'Increment
        RegCloseKey hCurKey 'Closes the current open key

        If apiError <> 234 Then
        If apiError > 0 Then
            Exit Do
        End If
        End If
    Loop While Not apiError = ERROR_NO_MORE_ITEMS

    If Not apiError = ERROR_NO_MORE_ITEMS Then
        Errors.Errors apiError, "RegEnumValue"
    End If
End Function

Public Function GetSettingBinary(hKey As Long, strPath As String, strValue As String) As Long
    Dim hCurKey As Long
    Dim lngBuffer As Long
    Dim lngDataBufferSize As Long
    Dim lngValueType As Long

    apiError = RegOpenKeyEx(hKey, strPath, vbNull, KEY_QUERY_VALUE, hCurKey)
    If Not apiError = 0 Then
        Errors.Errors apiError, "RegOpenKeyEx"
        Exit Function
    End If
    
    lngDataBufferSize = 4 '4 bytes = 32 bits = long

    apiError = RegQueryValueEx(hCurKey, strValue & Chr$(0), 0&, lngValueType, lngBuffer, lngDataBufferSize)
    If Not apiError = 0 Then Errors.Errors apiError, "RegQueryValueEx"

    If lngValueType = REG_BINARY Then
        GetSettingBinary = lngBuffer
    End If

    apiError = RegCloseKey(hCurKey) 'Closes the current open key
    If Not apiError = 0 Then Errors.Errors apiError, "RegCloseKey"
End Function

Public Function SaveSettingBinary(hKey As Long, strPath As String, strValue As String, lData As Long, Optional retDisposition As Long)
    Dim retKey As Long
    
    Dim SECURITY_ATTRIBUTES As SECURITY_ATTRIBUTES
    
    apiError = RegCreateKeyEx(hKey, strPath & Chr$(0), 0, Chr$(0), REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SECURITY_ATTRIBUTES, retKey, retDisposition)
    If Not apiError = 0 Then
        Errors.Errors apiError, "RegCreateKeyEx"
        Exit Function
    End If
    
    apiError = RegSetValueEx(retKey, strValue, 0&, REG_BINARY, lData, 4)
    If Not apiError = 0 Then Errors.Errors apiError, "RegSetValueEx"
    
    apiError = RegCloseKey(retKey) 'Closes the current open key
    If Not apiError = 0 Then Errors.Errors apiError, "RegCloseKey"
End Function

Public Function GetSettingByte(hKey As Long, strPath As String, strValueName As String) As Variant
    Dim byBuffer() As Byte
    Dim hCurKey As Long
    Dim lngDataBufferSize As Long
    Dim lngValueType As Long

    'Open the key and get number of bytes
    apiError = RegOpenKeyEx(hKey, strPath, vbNull, KEY_QUERY_VALUE, hCurKey)
    If Not apiError = 0 Then
        Errors.Errors apiError, "RegOpenKeyEx"
        Exit Function
    End If
    
    apiError = RegQueryValueEx(hCurKey, strValueName & Chr$(0), 0&, lngValueType, ByVal 0&, lngDataBufferSize)
    If Not apiError = 0 Then Errors.Errors apiError, "RegQueryValueEx"

    If lngValueType = REG_BINARY Then
        'initialise buffers and retrieve value
        ReDim byBuffer(lngDataBufferSize - 1) As Byte
        
        apiError = RegQueryValueEx(hCurKey, strValueName & Chr$(0), 0&, lngValueType, byBuffer(0), lngDataBufferSize)
        If Not apiError = 0 Then Errors.Errors apiError, "RegQueryValueEx"
        
        GetSettingByte = byBuffer
    End If

    apiError = RegCloseKey(hCurKey) 'Closes the current open key
End Function

Public Function GetSettingLong(hKey As Long, strPath As String, strValue As String) As Long
    Dim hCurKey As Long
    Dim lngBuffer As Long
    Dim lngDataBufferSize As Long
    Dim lngValueType As Long

    apiError = RegOpenKeyEx(hKey, strPath, vbNull, KEY_QUERY_VALUE, hCurKey)
    If Not apiError = 0 Then
        Errors.Errors apiError, "RegOpenKeyEx"
        Exit Function
    End If
    
    lngDataBufferSize = 4 '4 bytes = 32 bits = long

    apiError = RegQueryValueEx(hCurKey, strValue & Chr$(0), 0&, lngValueType, lngBuffer, lngDataBufferSize)
    If Not apiError = 0 Then Errors.Errors apiError, "RegQueryValueEx"
    
    If lngValueType = REG_DWORD Then
        GetSettingLong = lngBuffer
    End If
    
    apiError = RegCloseKey(hCurKey) 'Closes the current open key
    If Not apiError = 0 Then Errors.Errors apiError, "RegCloseKey"
End Function

Public Function GetSettingString(hKey As Long, strPath As String, strValue As String) As String
    Dim hCurKey As Long
    Dim lngDataBufferSize As Long
    Dim lngValueType As Long
    Dim strBuffer As String
    
    ' Open the key and get length of string
    apiError = RegOpenKeyEx(hKey, strPath, vbNull, KEY_QUERY_VALUE, hCurKey)
    If Not apiError = 0 Then
        Errors.Errors apiError, "RegOpenKeyEx"
        Exit Function
    End If
    
    apiError = RegQueryValueEx(hCurKey, strValue & Chr$(0), 0&, lngValueType, ByVal 0&, lngDataBufferSize)
    If Not apiError = 0 Then Errors.Errors apiError, "RegQueryValueEx"
    
    If lngValueType = REG_SZ Then
        strBuffer = Space$(lngDataBufferSize) 'Padding
        
        apiError = RegQueryValueEx(hCurKey, strValue & Chr$(0), 0&, 0&, ByVal strBuffer, lngDataBufferSize)
        If Not apiError = 0 Then Errors.Errors apiError, "RegQueryValueEx"
        
        GetSettingString = Fix_NullTermStr(strBuffer)
    End If

    apiError = RegCloseKey(hCurKey) 'Closes the current open key
    If Not apiError = 0 Then Errors.Errors apiError, "RegCloseKey"
End Function

Public Sub SaveSettingByte(hKey As Long, strPath As String, strValueName As String, byData() As Byte, Optional retDisposition As Long)
    Dim retKey As Long
    'Make sure that the array starts with element 0 before passing it! (otherwise it will not be saved!)
    
    Dim SECURITY_ATTRIBUTES As SECURITY_ATTRIBUTES
    
    apiError = RegCreateKeyEx(hKey, strPath & Chr$(0), 0, Chr$(0), REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SECURITY_ATTRIBUTES, retKey, retDisposition)
    If Not apiError = 0 Then
        Errors.Errors apiError, "RegCreateKeyEx"
        Exit Sub 'Cant open so exit
    End If
    
    'Pass the first array element and length of array
    apiError = RegSetValueEx(retKey, strValueName, 0&, REG_BINARY, byData(0), UBound(byData()) + 1)
    If Not apiError = 0 Then Errors.Errors apiError, "RegSetValueEx"
    
    apiError = RegCloseKey(retKey) 'Closes the current open key
    If Not apiError = 0 Then Errors.Errors apiError, "RegCloseKey"
End Sub

Public Sub SaveSettingLong(hKey As Long, strPath As String, strValue As String, lData As Long, Optional retDisposition As Long)
    Dim retKey As Long
    
    Dim SECURITY_ATTRIBUTES As SECURITY_ATTRIBUTES
    
    apiError = RegCreateKeyEx(hKey, strPath & Chr$(0), 0, Chr$(0), REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SECURITY_ATTRIBUTES, retKey, retDisposition)
    If Not apiError = 0 Then
        Errors.Errors apiError, "RegCreateKeyEx"
        Exit Sub 'Cant open so exit
    End If
    
    apiError = RegSetValueEx(retKey, strValue, 0&, REG_DWORD, lData, 4)
    If Not apiError = 0 Then Errors.Errors apiError, "RegSetValueEx"
    
    apiError = RegCloseKey(retKey) 'Closes the current open key
    If Not apiError = 0 Then Errors.Errors apiError, "RegCloseKey"
End Sub

Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String, Optional retDisposition As Long)
    Dim retKey As Long
    
    Dim SECURITY_ATTRIBUTES As SECURITY_ATTRIBUTES
    
    apiError = RegCreateKeyEx(hKey, strPath & Chr$(0), 0, Chr$(0), REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SECURITY_ATTRIBUTES, retKey, retDisposition)
    If Not apiError = 0 Then
        Errors.Errors apiError, "RegCreateKeyEx"
        Exit Sub 'Cant open so exit
    End If
    
    apiError = RegSetValueEx(retKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    If Not apiError = 0 Then Errors.Errors apiError, "RegSetValueEx"
    
    apiError = RegCloseKey(retKey) 'Closes the current open key
    If Not apiError = 0 Then Errors.Errors apiError, "RegCloseKey"
End Sub
