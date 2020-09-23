VERSION 5.00
Begin VB.Form frmExitWindows 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exit Windows"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmExitWindows.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkForceIfHung 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkForce 
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.ComboBox cboMethod 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   350
      Left            =   3120
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblForceIfHung 
      Caption         =   "Force If Hung"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblForce 
      Caption         =   "Force"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblMethod 
      Caption         =   "Method"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmExitWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkForce_Click()
    If chkForceIfHung.Value = 1 Then chkForceIfHung.Value = 0
End Sub

Private Sub chkForceIfHung_Click()
    If chkForce.Value = 1 Then chkForce.Value = 0
End Sub

Private Sub cmdOk_Click()
    Dim flags As Long
    
    Select Case cboMethod.ListIndex
        Case 0: flags = EWX_LOGOFF
        Case 1: flags = EWX_POWEROFF
        Case 2: flags = EWX_REBOOT
        Case 3: flags = EWX_SHUTDOWN
        Case Else: Exit Sub 'Nothing was selected
    End Select
    
    If chkForce.Value = 1 Then
        flags = flags Or EWX_FORCE
    End If
    If chkForceIfHung.Value = 1 Then
        flags = flags Or EWX_FORCEIFHUNG
    End If
    
    If WinID = "WIN32_WINDOWS" Then
        If ExitWindowsEx(flags, 0) = 0 Then Failed "ExitWindowsEx"
    Else 'NT
        'Adjust token
        
        Dim hCurProcess As Long
        Dim hTokenHandle As Long
        Dim tmpLuid As LUID
        Dim tkpNewState As TOKEN_PRIVILEGES
        Dim tkpPreviousState As TOKEN_PRIVILEGES
        Dim lngBufferLen As Long
         
        hCurProcess = GetCurrentProcess()
        
        If OpenProcessToken(hCurProcess, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hTokenHandle) = 0 Then Failed "OpenProcessToken"
        If LookupPrivilegeValue("", "SeShutdownPrivilege", tmpLuid) = 0 Then Failed "LookupPrivilegeValue"
        
        With tkpNewState
            .PrivilegeCount = 1
            .Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
            .Privileges(0).pLuid = tmpLuid
        End With
        
        If AdjustTokenPrivileges(hTokenHandle, False, tkpNewState, Len(tkpPreviousState), tkpPreviousState, lngBufferLen) = 0 Then Failed "LookupPrivilegeValue"
        If ExitWindowsEx(flags, 0) = 0 Then Failed "ExitWindowsEx"
    End If
End Sub

Private Sub Form_Load()
    With cboMethod
        .AddItem "Logoff"
        .AddItem "Poweroff"
        .AddItem "Reboot"
        .AddItem "Shutdown"
    End With
    
    If WinVersion(-1, 5000000) = True Then
        chkForceIfHung.Enabled = True
    End If
    
    chkForce.Value = GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira\ExitWindows", "Force")
    chkForceIfHung.Value = GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira\ExitWindows", "ForceIfHung")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\ExitWindows", "Force", chkForce.Value
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\ExitWindows", "ForceIfHung", chkForceIfHung.Value
    
    Unload Me
End Sub
