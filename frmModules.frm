VERSION 5.00
Begin VB.Form frmModules 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modules"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmModules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMUsageCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtMModule 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtMModuleHandle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtMGlobalUsageCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtMExePath 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtMBaseSize 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtMBaseAddress 
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
   Begin VB.ListBox lstModule 
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
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   4800
      TabIndex        =   20
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblMModuleHandle 
      Caption         =   "Module Handle"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblPExeFile 
      Caption         =   "Exe File"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblMExePath 
      Caption         =   "Exe Path"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblMBaseSize 
      Caption         =   "Base Size"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblMBaseAddress 
      Caption         =   "Base Address"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblMGlobalUsageCount 
      Caption         =   "Global Usage Count"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblMUsageCount 
      Caption         =   "Usage Count"
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblMModule 
      Caption         =   "Module Name"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblProcess 
      Caption         =   "Process"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblModule 
      Caption         =   "Module"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "frmModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Process() As PROCESSENTRY32
Dim lngProcess As Long
Dim Module() As MODULEENTRY32
Dim lngModule As Long

Private Sub cmdRefresh_Click()
    txtPExeFile.Text = ""
    
    txtMBaseAddress.Text = ""
    txtMBaseSize.Text = ""
    txtMExePath.Text = ""
    txtMGlobalUsageCount.Text = ""
    txtMModule.Text = ""
    txtMModuleHandle.Text = ""
    txtMUsageCount.Text = ""
    
    lstModule.Clear
    lngModule = 0
    ReDim Module(0)
    
    Form_Load
End Sub

Private Sub Form_Load()
    lstProcess.Clear
    lngProcess = 0
    ReDim Process(0)
    
    txtMBaseAddress.Text = ""
    txtMBaseSize.Text = ""
    txtMExePath.Text = ""
    txtMGlobalUsageCount.Text = ""
    txtMModule.Text = ""
    txtMModuleHandle.Text = ""
    txtMUsageCount.Text = ""
    
    
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

Private Sub lstModule_Click()
    txtMBaseAddress.Text = Module(lstModule.ListIndex + 1).modBaseAddr
    txtMBaseSize.Text = Module(lstModule.ListIndex + 1).modBaseSize
    txtMExePath.Text = Module(lstModule.ListIndex + 1).szExePath
    txtMGlobalUsageCount.Text = Module(lstModule.ListIndex + 1).GlblcntUsage
    txtMModule.Text = Module(lstModule.ListIndex + 1).szModule
    txtMModuleHandle.Text = Module(lstModule.ListIndex + 1).hModule
    txtMUsageCount.Text = Module(lstModule.ListIndex + 1).ProccntUsage
End Sub

Private Sub lstProcess_Click()
    txtPExeFile.Text = Process(lstProcess.ListIndex + 1).szExeFile
    
    
    Dim hSnapShot As Long
    
    lstModule.Clear
    lngModule = 0
    ReDim Module(0)
    
    apiError = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, Process(lstProcess.ListIndex + 1).th32ProcessID)
    If apiError = -1 Then
        Failed "CreateToolhelp32Snapshot"
        Exit Sub
    Else
        hSnapShot = apiError
    End If
    
    MODULEENTRY32.dwSize = Len(MODULEENTRY32)
    If Module32First(hSnapShot, MODULEENTRY32) = False Then
        Failed "Module32First"
        
        'Clean up
        If CloseHandle(hSnapShot) = 0 Then Failed "CloseHandle"
        Exit Sub
    Else
        lngModule = lngModule + 1
        ReDim Preserve Module(lngModule)
        Module(lngModule) = MODULEENTRY32
        
        lstModule.AddItem MODULEENTRY32.szModule
    End If
    
    Do
        'Cycle through until error
        If Module32Next(hSnapShot, MODULEENTRY32) = False Then
            Exit Do
        Else
            lngModule = lngModule + 1
            ReDim Preserve Module(lngModule)
            Module(lngModule) = MODULEENTRY32
                
            lstModule.AddItem MODULEENTRY32.szModule
        End If
    Loop

    If CloseHandle(hSnapShot) = 0 Then Failed "CloseHandle"
End Sub
