Attribute VB_Name = "winbase"
Option Explicit


Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function DebugActiveProcess Lib "kernel32" (ByVal dwProcessId As Long) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetProcessVersion Lib "kernel32" (ByVal ProcessId As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OsVersionInfo) As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function IsDebuggerPresent Lib "kernel32" () As Boolean
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Public Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Public Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long


    Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
    End Type

    Public Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
    End Type

    Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
    End Type

    Public Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
    End Type

    Public Type SYSTEM_POWER_STATUS
        ACLineStatus As Byte
        BatteryFlag As Byte
        BatteryLifePercent As Byte
        Reserved1 As Byte
        BatteryLifeTime As Long
        BatteryFullLifeTime As Long
    End Type
    
    Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
    End Type


    Public Const CREATE_NEW = 1
    Public Const CREATE_ALWAYS = 2
    Public Const OPEN_EXISTING = 3
    Public Const OPEN_ALWAYS = 4
    Public Const TRUNCATE_EXISTING = 5

    Public Const DRIVE_UNKNOWN = 0
    Public Const DRIVE_NO_ROOT_DIR = 1
    Public Const DRIVE_REMOVABLE = 2
    Public Const DRIVE_FIXED = 3
    Public Const DRIVE_REMOTE = 4
    Public Const DRIVE_CDROM = 5
    Public Const DRIVE_RAMDISK = 6

    Public Const FILE_FLAG_WRITE_THROUGH = &H80000000
    Public Const FILE_FLAG_OVERLAPPED = &H40000000
    Public Const FILE_FLAG_NO_BUFFERING = &H20000000
    Public Const FILE_FLAG_RANDOM_ACCESS = &H10000000
    Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
    Public Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
    Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
    Public Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
    Public Const FILE_FLAG_OPEN_REPARSE_POINT = &H200000
    Public Const FILE_FLAG_OPEN_NO_RECALL = &H100000

    Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
    Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
    Public Const FORMAT_MESSAGE_FROM_STRING = &H400
    Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
    Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
    Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
    Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

    Public Const FS_CASE_IS_PRESERVED = FILE_CASE_PRESERVED_NAMES
    Public Const FS_CASE_SENSITIVE = FILE_CASE_SENSITIVE_SEARCH
    Public Const FS_UNICODE_STORED_ON_DISK = FILE_UNICODE_ON_DISK
    Public Const FS_PERSISTENT_ACLS = FILE_PERSISTENT_ACLS
    Public Const FS_VOL_IS_COMPRESSED = FILE_VOLUME_IS_COMPRESSED
    Public Const FS_FILE_COMPRESSION = FILE_FILE_COMPRESSION
    Public Const FS_FILE_ENCRYPTION = FILE_SUPPORTS_ENCRYPTION

    Public Const MAX_COMPUTERNAME_LENGTH = 31

    Public Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
    Public Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
    Public Const THREAD_PRIORITY_NORMAL = 0
    Public Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
    Public Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
    Public Const THREAD_PRIORITY_ERROR_RETURN = (MAXLONG)
    Public Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
    Public Const THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE

Public Function Get_ComputerName() As String
    Dim strComputerName As String
    strComputerName = Space$(MAX_COMPUTERNAME_LENGTH + 1) 'Padd
    
    If GetComputerName(strComputerName, MAX_COMPUTERNAME_LENGTH + 1) = 0 Then
        Failed "GetComputerName"
    Else
        Get_ComputerName = Fix_NullTermStr(strComputerName)
    End If
End Function

Public Function Get_CurrentDirectory() As String
    Dim strCurrentDirectory As String
    strCurrentDirectory = Space$(1024) 'Padd
    
    If GetCurrentDirectory(1024, strCurrentDirectory) = 0 Then
        Failed "GetCurrentDirectory"
    Else
        Get_CurrentDirectory = Fix_Dir(Fix_NullTermStr(strCurrentDirectory))
    End If
End Function

Public Function Get_SystemDirectory() As String
    Dim strSystemDirectory As String
    strSystemDirectory = Space$(MAX_PATH) 'Padd
    
    If GetSystemDirectory(strSystemDirectory, MAX_PATH) = 0 Then
        Failed "GetSystemDirectory"
    Else
        Get_SystemDirectory = Fix_Dir(Fix_NullTermStr(strSystemDirectory))
    End If
End Function

Public Function Get_TempPath() As String
    Dim strTempPath As String
    strTempPath = Space$(1024) 'Padd
    
    If GetTempPath(1024, strTempPath) = 0 Then
        Failed "GetTempPath"
    Else
        Get_TempPath = Fix_Dir(Fix_NullTermStr(strTempPath))
    End If
End Function

Public Function Get_UserName() As String
    Dim strUserName As String
    strUserName = Space$(256 + 1) 'UNLEN + 1
    
    If GetUserName(strUserName, 256 + 1) = 0 Then
        Failed "GetUserName"
    Else
        Get_UserName = Fix_NullTermStr(strUserName) 'Send back out
    End If
End Function

Public Function Get_WindowsDirectory() As String
    Dim strWindowsDirectory As String
    strWindowsDirectory = Space$(MAX_PATH)  'Padd
    
    If GetWindowsDirectory(strWindowsDirectory, MAX_PATH) = 0 Then
        Failed "GetWindowsDirectory"
    Else
        Get_WindowsDirectory = Fix_Dir(Fix_NullTermStr(strWindowsDirectory))
    End If
End Function

Public Function Set_ComputerName(strName)
    If Len(strName) < 1 Then Exit Function 'Must be at least 1 in length
    
    'If to large trim the string
    If Len(strName) > MAX_COMPUTERNAME_LENGTH Then
        strName = Left$(strName, MAX_COMPUTERNAME_LENGTH)
    End If
    
    If apiError = SetComputerName(strName) Then Failed "SetComputerName"
End Function
