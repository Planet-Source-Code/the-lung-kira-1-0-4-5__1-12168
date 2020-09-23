VERSION 5.00
Begin VB.Form frmHeaps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Heaps"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frmHeaps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHLockCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtHHeapHandle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtHHFlags 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txtHBlockSize 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtHAddress 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ListBox lstHeaps 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   2535
   End
   Begin VB.TextBox txtHFlags 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   4800
      TabIndex        =   20
      Top             =   5640
      Width           =   975
   End
   Begin VB.ListBox lstHeap 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   2535
   End
   Begin VB.ListBox lstProcess 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txtPExeFile 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblHHFlags 
      Caption         =   "Flags"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblHLockCount 
      Caption         =   "Lock Count"
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label lblHBlockSize 
      Caption         =   "Block Size"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lblHAddress 
      Caption         =   "Address"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label lblHHeapHandle 
      Caption         =   "Heap Handle"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblHeaps 
      Caption         =   "Heap"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblHFlags 
      Caption         =   "Default Heap"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblHeap 
      Caption         =   "Heap List"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblProcess 
      Caption         =   "Process"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblPExeFile 
      Caption         =   "Exe File"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmHeaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Process() As PROCESSENTRY32
Dim lngProcess As Long
Dim Heap() As HEAPLIST32
Dim lngHeap As Long
Dim Heaps() As HEAPENTRY32
Dim lngHeaps As Long

Private Sub cmdRefresh_Click()
    txtPExeFile.Text = ""
    
    txtHFlags.Text = ""
    
    txtHAddress.Text = ""
    txtHBlockSize.Text = ""
    txtHHFlags.Text = ""
    txtHHeapHandle.Text = ""
    txtHLockCount.Text = ""
    
    lstHeap.Clear
    lstHeaps.Clear
    lngHeap = 0
    lngHeaps = 0
    ReDim Heap(0)
    ReDim Heaps(0)
    
    Form_Load
End Sub

Private Sub Form_Load()
    lstProcess.Clear
    lngProcess = 0
    ReDim Process(0)
    
    txtHFlags.Text = ""
    
    txtHAddress.Text = ""
    txtHBlockSize.Text = ""
    txtHHFlags.Text = ""
    txtHHeapHandle.Text = ""
    txtHLockCount.Text = ""
    
    
    If WinVersion(0, 5000000) = True Then
        Dim hSnapShot As Long
        
        apiError = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        If apiError = -1 Then
            Failed "CreateToolhelp32Snapshot"
            Exit Sub
        Else
            hSnapShot = apiError
        End If
        
        PROCESSENTRY32.dwSize = Len(PROCESSENTRY32)
        If Process32First(hSnapShot, PROCESSENTRY32) = False Then
            Failed "Process32First"
            
            'Clean up
            If CloseHandle(hSnapShot) = 0 Then Failed "CloseHandle"
            Exit Sub
        Else
            lngProcess = lngProcess + 1
            ReDim Preserve Process(lngProcess)
            Process(lngProcess) = PROCESSENTRY32
            
            lstProcess.AddItem PROCESSENTRY32.th32ProcessID
        End If

        Do
            'Cycle through until error
            If Process32Next(hSnapShot, PROCESSENTRY32) = False Then
                Exit Do
            Else
                lngProcess = lngProcess + 1
                ReDim Preserve Process(lngProcess)
                Process(lngProcess) = PROCESSENTRY32
                
                lstProcess.AddItem PROCESSENTRY32.th32ProcessID
            End If
        Loop
        
        If CloseHandle(hSnapShot) = 0 Then Failed "CloseHandle"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstHeap_Click()
    If Heap(lstHeap.ListIndex + 1).dwFlags = HF32_DEFAULT Then
        txtHFlags.Text = "True"
    Else
        txtHFlags.Text = "False"
    End If
    
    
    lstHeaps.Clear
    lngHeaps = 0
    ReDim Heaps(0)
    
    txtHAddress.Text = ""
    txtHBlockSize.Text = ""
    txtHHFlags.Text = ""
    txtHHeapHandle.Text = ""
    txtHLockCount.Text = ""
    
    
    HEAPENTRY32.dwSize = Len(HEAPENTRY32)
    If Heap32First(HEAPENTRY32, Process(lstProcess.ListIndex + 1).th32ProcessID, Heap(lstHeap.ListIndex + 1).th32HeapID) = False Then
        Failed "Heap32First"
        Exit Sub
    Else
        lngHeaps = lngHeaps + 1
        ReDim Preserve Heaps(lngHeaps)
        Heaps(lngHeaps) = HEAPENTRY32
        
        lstHeaps.AddItem HEAPENTRY32.th32HeapID
    End If
    
    Do
        'Cycle through until error
        If Heap32Next(HEAPENTRY32) = False Then
            Failed "Heap32Next"
            Exit Do
        Else
            lngHeaps = lngHeaps + 1
            ReDim Preserve Heaps(lngHeaps)
            Heaps(lngHeaps) = HEAPENTRY32
                
            lstHeaps.AddItem HEAPENTRY32.dwAddress
        End If
    Loop
End Sub

Private Sub lstHeaps_Click()
    txtHAddress.Text = Heaps(lstHeaps.ListIndex + 1).dwAddress
    txtHBlockSize.Text = Heaps(lstHeaps.ListIndex + 1).dwBlockSize
    
    Select Case Heaps(lstHeaps.ListIndex + 1).dwFlags
        Case LF32_FIXED: txtHHFlags.Text = "Fixed"
        Case LF32_FREE: txtHHFlags.Text = "Free"
        Case LF32_MOVEABLE: txtHHFlags.Text = "Moveable"
    End Select
    
    txtHHeapHandle.Text = Heaps(lstHeaps.ListIndex + 1).hHandle
    txtHLockCount.Text = Heaps(lstHeaps.ListIndex + 1).dwLockCount
End Sub

Private Sub lstProcess_Click()
    txtPExeFile.Text = Process(lstProcess.ListIndex + 1).szExeFile
    
    
    Dim hSnapShot As Long
    
    lstHeap.Clear
    lstHeaps.Clear
    lngHeap = 0
    lngHeaps = 0
    ReDim Heap(0)
    ReDim Heaps(0)
    
    apiError = CreateToolhelp32Snapshot(TH32CS_SNAPHEAPLIST, Process(lstProcess.ListIndex + 1).th32ProcessID)
    If apiError = -1 Then
        Failed "CreateToolhelp32Snapshot"
        Exit Sub
    Else
        hSnapShot = apiError
    End If
    
    HEAPLIST32.dwSize = Len(HEAPLIST32)
    If Heap32ListFirst(hSnapShot, HEAPLIST32) = False Then
        Failed "Heap32ListFirst"
        
        'Clean up
        If CloseHandle(hSnapShot) = 0 Then Failed "CloseHandle"
        Exit Sub
    Else
        lngHeap = lngHeap + 1
        ReDim Preserve Heap(lngHeap)
        Heap(lngHeap) = HEAPLIST32
        
        lstHeap.AddItem HEAPLIST32.th32HeapID
    End If
    
    Do
        'Cycle through until error
        If Heap32ListNext(hSnapShot, HEAPLIST32) = False Then
            Exit Do
        Else
            lngHeap = lngHeap + 1
            ReDim Preserve Heap(lngHeap)
            Heap(lngHeap) = HEAPLIST32
                
            lstHeap.AddItem HEAPLIST32.th32HeapID
        End If
    Loop

    If CloseHandle(hSnapShot) = 0 Then Failed "CloseHandle"
End Sub
