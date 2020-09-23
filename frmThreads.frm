VERSION 5.00
Begin VB.Form frmThreads 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Threads"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmThreads.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2265
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
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtThreadID 
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
   Begin VB.TextBox txtOwnerProcessID 
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
   Begin VB.TextBox txtDeltaPriority 
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
   Begin VB.TextBox txtBasePriority 
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
      TabIndex        =   12
      Top             =   1800
      Width           =   975
   End
   Begin VB.ListBox lstThread 
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
   Begin VB.Label lblDeltaPriority 
      Caption         =   "Delta Priority"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblBasePriority 
      Caption         =   "Base Priority"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblOwnerProcessID 
      Caption         =   "Owner Process ID"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblThreadID 
      Caption         =   "Thread ID"
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
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblThread 
      Caption         =   "Thread"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmThreads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Thread() As THREADENTRY32
Dim lngThread As Long

Private Sub cmdRefresh_Click()
    txtBasePriority.Text = ""
    txtDeltaPriority.Text = ""
    txtOwnerProcessID.Text = ""
    txtThreadID.Text = ""
    txtUsage.Text = ""
    
    Form_Load
End Sub

Private Sub Form_Load()
    If WinVersion(0, 5000000) = True Then
        Dim hSnapShot As Long
        
        lstThread.Clear
        lngThread = 0
        ReDim Thread(0)
        
        apiError = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0&)
        If apiError = -1 Then
            Failed "CreateToolhelp32Snapshot"
            Exit Sub
        Else
            hSnapShot = apiError
        End If
        
        THREADENTRY32.dwSize = Len(THREADENTRY32)
        If Thread32First(hSnapShot, THREADENTRY32) = False Then
            Failed "Thread32First"
            
            'Clean up
            If CloseHandle(hSnapShot) = 0 Then Failed "CloseHandle"
            Exit Sub
        Else
            lngThread = lngThread + 1
            ReDim Preserve Thread(lngThread)
            Thread(lngThread) = THREADENTRY32
            
            lstThread.AddItem THREADENTRY32.th32ThreadID
        End If
        
        Do
            'Cycle through until error
            If Thread32Next(hSnapShot, THREADENTRY32) = False Then
                Exit Do
            Else
                lngThread = lngThread + 1
                ReDim Preserve Thread(lngThread)
                Thread(lngThread) = THREADENTRY32
                    
                lstThread.AddItem THREADENTRY32.th32ThreadID
            End If
        Loop
    
        If CloseHandle(hSnapShot) = 0 Then Failed "CloseHandle"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstThread_Click()
    Select Case Thread(lstThread.ListIndex + 1).tpBasePri
        Case THREAD_PRIORITY_LOWEST: txtBasePriority.Text = "Lowest"
        Case THREAD_PRIORITY_BELOW_NORMAL: txtBasePriority.Text = "Below Normal"
        Case THREAD_PRIORITY_NORMAL: txtBasePriority.Text = "Normal"
        Case THREAD_PRIORITY_HIGHEST: txtBasePriority.Text = "Highest"
        Case THREAD_PRIORITY_ABOVE_NORMAL: txtBasePriority.Text = "Above Normal"
        Case THREAD_PRIORITY_ERROR_RETURN: txtBasePriority.Text = "Error Return"
        Case THREAD_PRIORITY_TIME_CRITICAL: txtBasePriority.Text = "Time Critical"
        Case THREAD_PRIORITY_IDLE: txtBasePriority.Text = "Idle"
    End Select
    
    txtDeltaPriority.Text = Thread(lstThread.ListIndex + 1).tpDeltaPri
    txtOwnerProcessID.Text = Thread(lstThread.ListIndex + 1).th32OwnerProcessID
    txtThreadID.Text = Thread(lstThread.ListIndex + 1).th32ThreadID
    txtUsage.Text = Thread(lstThread.ListIndex + 1).cntUsage
End Sub
