VERSION 5.00
Begin VB.Form frmStartMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Start Menu"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "frmStartMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRun 
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkFind 
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkRecentDocsMenu 
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkRecentDocsHistory 
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkLogoff 
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2400
      TabIndex        =   12
      Top             =   1680
      Width           =   975
   End
   Begin VB.CheckBox chkFavoritesMenu 
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblRun 
      Caption         =   "Run"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   2805
   End
   Begin VB.Label lblFind 
      Caption         =   "Find"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2805
   End
   Begin VB.Label lblRecentDocsMenu 
      Caption         =   "Recent Docs Menu"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   2805
   End
   Begin VB.Label lblRecentDocsHistory 
      Caption         =   "Recent Docs History"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2805
   End
   Begin VB.Label lblLogoff 
      Caption         =   "Logoff"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2805
   End
   Begin VB.Label lblFavoritesMenu 
      Caption         =   "Favorites Menu"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2805
   End
End
Attribute VB_Name = "frmStartMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    If chkFavoritesMenu.Value = 0 Then
        SaveSettingBinary HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", 1
    Else
        SaveSettingBinary HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", 0
    End If
    If chkFind.Value = 0 Then
        SaveSettingBinary HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", 1
    Else
        SaveSettingBinary HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", 0
    End If
    If chkLogoff.Value = 0 Then
        SaveSettingBinary HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", 1
    Else
        SaveSettingBinary HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", 0
    End If
    If chkRecentDocsHistory.Value = 0 Then
        SaveSettingBinary HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsHistory", 1
    Else
        SaveSettingBinary HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsHistory", 0
    End If
    If chkRecentDocsMenu.Value = 0 Then
        SaveSettingBinary HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu", 1
    Else
        SaveSettingBinary HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu", 0
    End If
    If chkRun.Value = 0 Then
        SaveSettingBinary HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 1
    Else
        SaveSettingBinary HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 0
    End If
End Sub

Private Sub Form_Load()
    Select Case GetSettingBinary(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu")
        Case 0: chkFavoritesMenu.Value = 1
        Case 1: chkFavoritesMenu.Value = 0
    End Select
    Select Case GetSettingBinary(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind")
        Case 0: chkFind.Value = 1
        Case 1: chkFind.Value = 0
    End Select
    Select Case GetSettingBinary(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff")
        Case 0: chkLogoff.Value = 1
        Case 1: chkLogoff.Value = 0
    End Select
    Select Case GetSettingBinary(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsHistory")
        Case 0: chkRecentDocsHistory.Value = 1
        Case 1: chkRecentDocsHistory.Value = 0
    End Select
    Select Case GetSettingBinary(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu")
        Case 0: chkRecentDocsMenu.Value = 1
        Case 1: chkRecentDocsMenu.Value = 0
    End Select
    Select Case GetSettingBinary(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun")
        Case 0: chkRun.Value = 1
        Case 1: chkRun.Value = 0
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
