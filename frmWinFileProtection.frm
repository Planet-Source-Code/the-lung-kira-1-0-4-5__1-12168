VERSION 5.00
Begin VB.Form frmWinFileProtection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows File Protection"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   Icon            =   "frmWinFileProtection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFiles 
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
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   9255
   End
   Begin VB.TextBox txtFiles 
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
      TabIndex        =   0
      Top             =   2760
      Width           =   9255
   End
End
Attribute VB_Name = "frmWinFileProtection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim PROTECTED_FILE_DATA As PROTECTED_FILE_DATA
    
    If WinVersion(-1, 5000000) = True Then
        PROTECTED_FILE_DATA.FileNumber = 0
        If SfcGetNextProtectedFile(0&, PROTECTED_FILE_DATA) = 0 Then Failed "SfcGetNextProtectedFile"
        
        Dim lngIncrement As Long
        
        Do
            lngIncrement = lngIncrement + 1
            
            PROTECTED_FILE_DATA.FileNumber = lngIncrement
            If SfcGetNextProtectedFile(0&, PROTECTED_FILE_DATA) = 0 Then
                Failed "SfcGetNextProtectedFile"
                Exit Do
            End If
            
            lstFiles.AddItem Replace$(PROTECTED_FILE_DATA.FileName, Chr(0), "", 1, -1)
            PROTECTED_FILE_DATA.FileName = Space$(MAX_PATH)
        Loop
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstFiles_Click()
    txtFiles.Text = lstFiles.List(lstFiles.ListIndex)
End Sub
