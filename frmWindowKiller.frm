VERSION 5.00
Begin VB.Form frmWindowKiller 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Window Killer"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmWindowKiller.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImportList 
      Caption         =   "Import List"
      Height          =   350
      Left            =   5640
      TabIndex        =   9
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   350
      Left            =   4680
      TabIndex        =   8
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   350
      Left            =   3600
      TabIndex        =   7
      Top             =   4440
      Width           =   975
   End
   Begin VB.ListBox lstWindows 
      Height          =   1425
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   6495
   End
   Begin VB.TextBox txtKill 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   6495
   End
   Begin VB.ListBox lstKill 
      Height          =   1425
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   6495
   End
   Begin VB.CommandButton cmdRefreshWin 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   5640
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   350
      Left            =   2640
      TabIndex        =   6
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lblWindows 
      Caption         =   "Available Windows"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblWindowKiller 
      Caption         =   "Window Killer List"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
End
Attribute VB_Name = "frmWindowKiller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    If txtKill.Text <> "" Then 'Cant enter a blank entry
        lstKill.AddItem txtKill.Text 'Add item to list box
        txtKill.Text = ""
        Call Refresh_Array
    End If
End Sub

Private Sub cmdClearAll_Click()
    lstKill.Clear
    Call Refresh_Array
End Sub

Private Sub cmdImportList_Click()
    Dim strFileName As String
    GetOpenName hwnd, "Open", strFileName
    
    'Error checking
    If Not strFileName <> "" Then Exit Sub 'Dont worry just exit
    If Not FileLen(strFileName) > 0 Then Exit Sub 'If file len not greater than 0
    
    Dim strFileContents As String
    Open strFileName For Input As #1
        Do While Not EOF(1) 'Loop until end of file
            Line Input #1, strFileContents 'Read line into variable
            If strFileContents <> "" Then lstKill.AddItem strFileContents
        Loop
    Close #1
    
    Call Refresh_Array
End Sub

Private Sub cmdRefreshWin_Click()
    'Clear
    lstWindows.Clear
    ReDim WindowListName(0)
    ReDim WindowListhWnd(0)
    WindowListNum = 0
    
    'Enumerate all the handles
    If EnumWindows(AddressOf EnumWindowsProc, 0&) = 0 Then Failed "EnumWindows"

    Dim lngIncrement As Long
    For lngIncrement = 0 To WindowListNum - 1 'Cycle through list
        If WindowListName(lngIncrement) <> "" Then 'If text then add
            lstWindows.AddItem WindowListName(lngIncrement)
        End If
    Next lngIncrement
End Sub

Private Sub cmdRemove_Click()
    On Error Resume Next 'Rather skip the error than call the function 2x
    lstKill.RemoveItem lstKill.ListIndex
    Call Refresh_Array
End Sub

Private Sub Form_Load()
    If WindowKillerNum > 0 Then
        Dim lngIncrement As Long
    
        For lngIncrement = 1 To WindowKillerNum
            lstKill.AddItem WindowKiller(lngIncrement)
        Next lngIncrement
    End If
    
    Call cmdRefreshWin_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstKill_Click()
    txtKill.Text = lstKill.List(lstKill.ListIndex)
End Sub

Private Sub lstWindows_Click()
    txtKill.Text = lstWindows.List(lstWindows.ListIndex)
End Sub

Private Function Refresh_Array()
    If lstKill.ListCount = 0 Then
        Erase WindowKiller() 'Erase array if nothing
    Else
        Dim lngIncrement As Long
        
        For lngIncrement = 0 To lstKill.ListCount - 1
            ReDim Preserve WindowKiller(lngIncrement + 1)
            WindowKiller(lngIncrement + 1) = lstKill.List(lngIncrement)
        Next lngIncrement
        
        WindowKillerNum = lngIncrement
    End If
End Function
