VERSION 5.00
Begin VB.Form frmDirectories 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Directories"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   Icon            =   "frmDirectories.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDirectories 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Width           =   9255
   End
   Begin VB.ListBox lstDirectories 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmDirectories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Dim strBuffer As String * MAX_PATH
    
    With lstDirectories
        'If SHGetSpecialFolderPath(0&, strBuffer, CSIDL_ADMINTOOLS, False) = True Then
        '    .AddItem Left$("AdminTools" & Space$(15), 15) & Fix_NullTermStr(strBuffer)
        'Else
        '    Failed "SHGetSpecialFolderPath"
        'End If
        
        .AddItem Left$("AdminTools" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Administrative Tools")
        .AddItem Left$("AppData" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "AppData")
        .AddItem Left$("AppPath" & Space$(15), 15) & Dirs.AppPath
        .AddItem Left$("Cache" & Space$(15), 15) & Dirs.Cache
        .AddItem Left$("CommonFiles" & Space$(15), 15) & GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "CommonFilesDir")
        .AddItem Left$("Cookies" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Cookies")
        .AddItem Left$("Current" & Space$(15), 15) & Get_CurrentDirectory
        .AddItem Left$("Desktop" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop")
        .AddItem Left$("Favorites" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Favorites")
        .AddItem Left$("Fonts" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Fonts")
        .AddItem Left$("History" & Space$(15), 15) & Dirs.History
        .AddItem Left$("LocalAppData" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Local AppData")
        .AddItem Left$("MediaPath" & Space$(15), 15) & GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "MediaPath")
        .AddItem Left$("MyPictures" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "My Pictures")
        .AddItem Left$("NetHood" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "NetHood")
        .AddItem Left$("Personal" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal")
        .AddItem Left$("PrintHood" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "PrintHood")
        .AddItem Left$("Programs" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Programs")
        .AddItem Left$("Recent" & Space$(15), 15) & Dirs.Recent
        .AddItem Left$("SendTo" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "SendTo")
        .AddItem Left$("StartMenu" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Start Menu")
        .AddItem Left$("Startup" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup")
        .AddItem Left$("System" & Space$(15), 15) & Dirs.System
        .AddItem Left$("Temp" & Space$(15), 15) & Dirs.Temp
        .AddItem Left$("Templates" & Space$(15), 15) & GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Templates")
        .AddItem Left$("Windows" & Space$(15), 15) & Dirs.Windows
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstDirectories_Click()
    txtDirectories.Text = Right$(lstDirectories.List(lstDirectories.ListIndex), Len(lstDirectories.List(lstDirectories.ListIndex)) - 15)
End Sub
