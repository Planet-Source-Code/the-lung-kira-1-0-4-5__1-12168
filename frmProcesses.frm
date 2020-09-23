VERSION 5.00
Begin VB.Form frmProcesses 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Processes"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmProcesses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUsage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtThreads 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtProcessID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtPrimaryBaseClass 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtParentProcessID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtExeFile 
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
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   4800
      TabIndex        =   14
      Top             =   2040
      Width           =   975
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
   Begin VB.Label lblExeFile 
      Caption         =   "Exe File"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblPrimaryBaseClass 
      Caption         =   "Primary Base Class"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblParentProcessID 
      Caption         =   "Parent Process ID"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblThreads 
      Caption         =   "Threads"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblProcessID 
      Caption         =   "Process ID"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblUsage 
      Caption         =   "Usage"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   1560
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
End
Attribute VB_Name = "frmProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Process() As PROCESSENTRY32
Dim lngProcess As Long

Private Sub cmdRefresh_Click()
    txtExeFile.Text = ""
    txtParentProcessID.Text = ""
    txtPrimaryBaseClass.Text = ""
    txtProcessID.Text = ""
    txtThreads.Text = ""
    txtUsage.Text = ""
    
    Form_Load
End Sub

Private Sub Form_Load()
    lstProcess.Clear
    lngProcess = 0
    ReDim Process(0)
    
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

Private Sub lstProcess_Click()
    txtExeFile.Text = Process(lstProcess.ListIndex + 1).szExeFile
    txtParentProcessID.Text = Process(lstProcess.ListIndex + 1).th32ParentProcessID
    txtPrimaryBaseClass.Text = Process(lstProcess.ListIndex + 1).pcPriClassBase
    txtProcessID.Text = Process(lstProcess.ListIndex + 1).th32ProcessID
    txtThreads.Text = Process(lstProcess.ListIndex + 1).cntThreads
    txtUsage.Text = Process(lstProcess.ListIndex + 1).cntUsage
End Sub
