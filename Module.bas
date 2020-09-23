Attribute VB_Name = "Module"
Option Explicit

    Public apiError As Long
    Public errMsg As Boolean
    Public ComputerName As String
    Public UserName As String
    Public WinVer As Long
    Public WinID As String
    
    Public Dirs As Dirs
    Public Type Dirs
        AppPath As String
        Cache As String
        History As String
        Recent As String
        System As String
        Temp As String
        Windows As String
    End Type
    
    Public GUID As GUID
    Public Type GUID
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(8) As Byte
    End Type
    
    Public MpgInfo As MpgInfo
    Public Type MpgInfo
        Sync As String
        Version As String
        Layer As Byte
        Error_Protection As Integer 'Bool
        Bitrate_Index As Integer
        Sampling_Freq As Long
        Padding As String
        Extension As String
        Mode As String
        Mode_Extn As String
        Copyright As Integer 'Bool
        Original As Integer 'Bool
        Emphasis As String
    End Type
    
    Public MpgTag As MpgTag
    Public Type MpgTag
        Tag As Boolean
        Title As String * 30
        Artist As String * 30
        Album As String * 30
        Year As String * 4
        Comments As String * 30
        Genre As Byte
    End Type

    Public WinsockData As WinsockData
    Public Type WinsockData
        Description As String
        SystemStatus As String
    End Type
'--------------------------------------
    Public WindowListName() As String
    Public WindowListhWnd() As Long
    Public WindowListNum As Long
    Public MouseMovX As Double
    Public MouseMovY As Double
    Public MouseMovTmpX As Integer
    Public MouseMovTmpY As Integer
    Public MouseWarp As Double
    Public WindowKiller() As String
    Public WindowKillerNum As Integer
    
    Public ScreenEdge As ScreenEdge
    Public Type ScreenEdge
        X As Integer
        Y As Integer
    End Type

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim strWindowText As String * 512
    
    apiError = GetWindowText(hwnd, strWindowText, 512)
    If apiError = 0 Then Failed "GetWindowText"
    
    ReDim Preserve WindowListName(WindowListNum) 'Resize arrays
    ReDim Preserve WindowListhWnd(WindowListNum)
    WindowListName(WindowListNum) = Fix_NullTermStr(strWindowText)
    WindowListhWnd(WindowListNum) = hwnd
    
    WindowListNum = WindowListNum + 1 'Increment
    
    EnumWindowsProc = 1
End Function

Public Function App_Shutdown()
    WinSockEnd 'Shutdown winsock
    StopHooks
    
    'Save settings
    'Need a # not true false
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira", "MouseMovOO", CByte(frmMain.mnuMouseMovOO.Checked)
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira", "MouseWarpOO", CByte(frmMain.mnuMouseWarpOO.Checked)
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira", "WindowKiller", CByte(frmMain.mnuWindowKillerOO.Checked)
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira", "errMsg", CByte(errMsg)
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Kira\MouseMov", "MouseMovX", CStr(MouseMovX)
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Kira\MouseMov", "MouseMovY", CStr(MouseMovY)
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Kira\MouseWarp", "MouseWarp", CStr(MouseWarp)
    
    If WindowKillerNum > 0 Then 'Dont write no data
        Dim tmpInt As Integer
        Dim tmpArray() As String
        Dim tmpCount As Integer
    
        ReDim Preserve tmpArray(1)
        tmpCount = 1
        tmpArray(1) = WindowKiller(1)
        
        For tmpInt = 1 To WindowKillerNum 'Cycle through WindowKiller array
            If Not tmpArray(tmpCount) = WindowKiller(tmpInt) Then 'If not duplicate
                tmpCount = tmpCount + 1 'Increment
                ReDim Preserve tmpArray(tmpCount) 'Resize array
                tmpArray(tmpCount) = WindowKiller(tmpInt) 'Add item
            End If
        Next tmpInt
    
        Open Dirs.AppPath & "\wk.dat" For Output As #1
            For tmpInt = 1 To tmpCount 'Cycle through tmpArray
                Print #1, tmpArray(tmpInt) 'Put in file
            Next tmpInt
        Close #1
    End If
    
    'Remove icon from system tray
    Dim NOTIFYICONDATA As NOTIFYICONDATA
    With NOTIFYICONDATA
        .cbSize = Len(NOTIFYICONDATA)
        .hwnd = frmMain.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = vbNull
        .hIcon = frmMain.Icon
        .szTip = Chr$(0)    'Clear
    End With
    If Shell_NotifyIcon(NIM_DELETE, NOTIFYICONDATA) = 0 Then Failed "Shell_NotifyIcon"
    
    End 'Exit program
End Function

Public Function App_Startup()
    If App.PrevInstance = True Then End

    'Add icon to system tray
    Dim NOTIFYICONDATA As NOTIFYICONDATA
    With NOTIFYICONDATA
        .cbSize = Len(NOTIFYICONDATA)
        .hwnd = frmMain.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = 512
        .hIcon = frmMain.Icon
        .szTip = frmMain.Caption & Chr$(0) 'Tooltip text
    End With
    If Shell_NotifyIcon(NIM_ADD, NOTIFYICONDATA) = 0 Then Failed "Shell_NotifyIcon"
    
    'Startup Windows Info
    Dim OsVersionInfo As OsVersionInfo
    OsVersionInfo.dwOSVersionInfoSize = Len(OsVersionInfo) 'Size of the structure
    If GetVersionEx(OsVersionInfo) = 0 Then Failed "GetVersionEx"

    Select Case OsVersionInfo.dwPlatformId
        Case VER_PLATFORM_WIN32_NT: WinID = "WIN32_NT"
        Case VER_PLATFORM_WIN32_WINDOWS: WinID = "WIN32_WINDOWS"
        Case VER_PLATFORM_WIN32s: WinID = "WIN32s"
    End Select
    
    WinVer = Right$("0" & OsVersionInfo.dwMajorVersion, 1) & _
             Right$("00" & OsVersionInfo.dwMinorVersion, 2) & _
             Right$("0000" & (OsVersionInfo.dwBuildNumber And &HFFFF&), 4)
    
    With Dirs
        .AppPath = Fix_Dir(App.path)
        .Cache = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Cache"))
        .History = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "History"))
        .Recent = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Recent"))
        .System = Get_SystemDirectory
        .Temp = Get_TempPath
        .Windows = Get_WindowsDirectory
    End With
    
    ComputerName = Get_ComputerName
    UserName = Get_UserName
    If cpuid_avail = 1 Then cpu_id 'Initializes function
    WinSockStart 'Startup winsock 2.2
    'StartHooks
'--------------------------------------
    'Read settings
    If Not GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira", "Kira") = 1 Then
        DefaultSettings
    End If
    
    frmMain.mnuMouseMovOO.Checked = CBool(GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira", "MouseMovOO"))
    frmMain.mnuMouseWarpOO.Checked = CBool(GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira", "MouseWarpOO"))
    frmMain.mnuWindowKillerOO.Checked = CBool(GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira", "WindowKiller"))
    errMsg = CBool(GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira", "errMsg"))
    'frmMain.mnuErrMsg.Checked = errMsg
    MouseMovX = CDbl(GetSettingString(HKEY_LOCAL_MACHINE, "Software\Kira\MouseMov", "MouseMovX"))
    MouseMovY = CDbl(GetSettingString(HKEY_LOCAL_MACHINE, "Software\Kira\MouseMov", "MouseMovY"))
    MouseWarp = CDbl(GetSettingString(HKEY_LOCAL_MACHINE, "Software\Kira\MouseWarp", "MouseWarp"))
    
    If Dir$(Dirs.AppPath & "\wk.dat") <> "" Then
        Dim tmpLong As Long
        Dim tmpString As String
        
        Open Dirs.AppPath & "\wk.dat" For Input As #1
            Do While Not EOF(1) 'Loop until end of file
                Line Input #1, tmpString 'Read line into variable
                tmpLong = tmpLong + 1 'Increment
                
                ReDim Preserve WindowKiller(tmpLong)
                WindowKiller(tmpLong) = Trim$(tmpString)
            Loop
        Close #1
        
        WindowKillerNum = tmpLong 'Propriatary
    End If
'--------------------------------------
    'Get it going
    If frmMain.mnuMouseMovOO.Checked = True Then
        Dim POINTAPI As POINTAPI
        GetCursorPos POINTAPI 'Dumps info to pointapi
    
        'Gives point of reference for starting
        MouseMovTmpX = POINTAPI.X
        MouseMovTmpY = POINTAPI.Y
    End If
    If frmMain.mnuMouseWarpOO.Checked = True Then
        'Converts the twips to pixels and sets the edges
        ScreenEdge.X = Screen.Width \ Screen.TwipsPerPixelX
        ScreenEdge.Y = Screen.Height \ Screen.TwipsPerPixelY
    End If
End Function

Public Sub DefaultSettings()
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira", "Kira", 1
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira", "MouseMovOO", 0
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira", "MouseWarpOO", 0
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira", "WindowKiller", 0
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira", "errMsg", 0
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\CDPlayer", "FFRW", 1000
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\DayTime", "Method", 1
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\Discard", "Method", 1
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\DiskSpace", "Output", 4
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\DisplaySettings", "GlobalChange", 1
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\ExitWindows", "Force", 0
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\ExitWindows", "ForceIfHung", 0
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\IEHistory", "Output", 0
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\IEHistory", "Sorted", 1
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\MemoryStatus", "Output", 3
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Kira\MouseMov", "MouseMovX", "0"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Kira\MouseMov", "MouseMovY", "0"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Kira\MouseWarp", "MouseWarp", "0"
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\PerfMon", "Interval", 1000
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\Ping", "Number", 1
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\Ping", "Timeout", 5000
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\Ping", "TTL", 255
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\Time", "Method", 1
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\WindowKiller", "Interval", 200
End Sub
