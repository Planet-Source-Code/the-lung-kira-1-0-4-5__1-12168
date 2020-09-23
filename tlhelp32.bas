Attribute VB_Name = "tlhelp32"
Option Explicit


Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Heap32ListFirst Lib "kernel32" (ByVal hSnapShot As Long, lphl As HEAPLIST32) As Long
Public Declare Function Heap32ListNext Lib "kernel32" (ByVal hSnapShot As Long, lphl As HEAPLIST32) As Long
Public Declare Function Heap32First Lib "kernel32" (lphe As HEAPENTRY32, ByVal th32ProcessID As Long, ByVal th32HeapID As Long) As Long
Public Declare Function Heap32Next Lib "kernel32" (lphe As HEAPENTRY32) As Long
Public Declare Function Module32First Lib "kernel32" (ByVal hSnapShot As Long, lpme As MODULEENTRY32) As Long
Public Declare Function Module32Next Lib "kernel32" (ByVal hSnapShot As Long, lpme As MODULEENTRY32) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Thread32First Lib "kernel32" (ByVal hSnapShot As Long, lpte As THREADENTRY32) As Long
Public Declare Function Thread32Next Lib "kernel32" (ByVal hSnapShot As Long, lpte As THREADENTRY32) As Long
'Public Declare Function Toolhelp32ReadProcessMemory Lib "kernel32" (ByVal th32ProcessID As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal cbRead As Long, ByVal lpNumberOfBytesRead As Long) As Boolean


    Public HEAPENTRY32 As HEAPENTRY32
    Public Type HEAPENTRY32
        dwSize As Long
        hHandle As Long         'Handle of this heap block
        dwAddress As Long       'Linear address of start of block
        dwBlockSize As Long     'Size of block in bytes
        dwFlags As Long
        dwLockCount As Long
        dwResvd As Long
        th32ProcessID As Long   'owning process
        th32HeapID As Long      'heap block is in
    End Type
    
    Public HEAPLIST32 As HEAPLIST32
    Public Type HEAPLIST32
        dwSize As Long
        th32ProcessID As Long   'owning process
        th32HeapID As Long      'heap (in owning process's context!)
        dwFlags As Long
    End Type

    Public MODULEENTRY32 As MODULEENTRY32
    Public Type MODULEENTRY32
        dwSize As Long
        th32ModuleID As Long        'This module
        th32ProcessID As Long       'owning process
        GlblcntUsage As Long        'Global usage count on the module
        ProccntUsage As Long        'Module usage count in th32ProcessID's context
        modBaseAddr As Long         'Base address of module in th32ProcessID's context
        modBaseSize As Long         'Size in bytes of module starting at modBaseAddr
        hModule As Long             'The hModule of this module in th32ProcessID's context
        szModule As String * 256    'MAX_MODULE_NAME32 + 1
        szExePath As String * MAX_PATH
    End Type

    Public PROCESSENTRY32 As PROCESSENTRY32
    Public Type PROCESSENTRY32
        dwSize As Long
        cntUsage As Long
        th32ProcessID  As Long          'this process
        th32DefaultHeapID As Long
        th32ModuleID As Long            'associated exe
        cntThreads As Long
        th32ParentProcessID As Long     'this process's parent process
        pcPriClassBase As Long          'Base priority of process's threads
        dwFlags As Long
        szExeFile As String * MAX_PATH  'Path
    End Type

    Public THREADENTRY32 As THREADENTRY32
    Public Type THREADENTRY32
        dwSize As Long
        cntUsage As Long
        th32ThreadID As Long        'this thread
        th32OwnerProcessID As Long  'Process this thread is associated with
        tpBasePri As Long
        tpDeltaPri As Long
        dwFlags As Long
    End Type

    
    Public Const HF32_DEFAULT = 1   'process's default heap
    Public Const HF32_SHARED = 2    'is shared heap

    Public Const LF32_FIXED = &H1
    Public Const LF32_FREE = &H2
    Public Const LF32_MOVEABLE = &H4

    Public Const MAX_MODULE_NAME32 = 255

    Public Const TH32CS_SNAPHEAPLIST = &H1
    Public Const TH32CS_SNAPPROCESS = &H2
    Public Const TH32CS_SNAPTHREAD = &H4
    Public Const TH32CS_SNAPMODULE = &H8
    Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
    Public Const TH32CS_INHERIT = &H80000000
