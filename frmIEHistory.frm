VERSION 5.00
Begin VB.Form frmIEHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IE History Viewer"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   Icon            =   "frmIEHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboOutput 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.CheckBox chkSorted 
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtFileSize 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   340
      Left            =   1800
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblOutput 
      Caption         =   "Output"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblSorted 
      Caption         =   "Sorted"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblTotal 
      Caption         =   "Total Links"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblFileSize 
      Caption         =   "File Size"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmIEHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
    'Clear
    txtTotal.Text = ""
    txtFileSize.Text = ""
    
    If Dir$(Dirs.History & "\History.IE5\index.dat") = "" Then
        MsgBox "Missing " & vbCrLf & Dirs.History & "\History.IE5\index.dat", vbInformation, "Error"
        Exit Sub 'Cant continue
    End If
    
    Dim strFileContents As String
    Dim StartPos As Long, EndPos As Long
    Dim TotalLinks As Long
    Dim AddLink As String
    Dim tmpAray() As String 'Dynamic array

    txtFileSize.Text = Round((FileLen(Dirs.History & "\History.IE5\index.dat") / 1024 ^ 2), 3) & " MB" 'Gives file size to text box"
    
    Open Dirs.History & "\History.IE5\index.dat" For Binary As #1 'Opens it for binary
        strFileContents = Space$(LOF(1)) 'Pads to length of string
        Get #1, , strFileContents 'Dumps contents of file to string
    Close #1
    
    'Cant allow zero in the search string
    StartPos = 1
    EndPos = 1
    
    Do
        StartPos = InStr(EndPos, strFileContents, "Visited: " & UserName & "@") 'Searches for visited string
        If StartPos = 0 Then Exit Do 'If not found then exit loop
        
        EndPos = InStr(StartPos, strFileContents, Chr$(0)) 'Searches for null terminator
        
        TotalLinks = TotalLinks + 1 'Incriment
        txtTotal.Text = TotalLinks 'Dump to text box
        ReDim Preserve tmpAray(TotalLinks) 'Resizes array without destroying
        
        AddLink = Mid$(strFileContents, StartPos + (Len(UserName) + 10), EndPos - (StartPos + (Len(UserName) + 10))) 'Pulls out url
        tmpAray(TotalLinks) = AddLink 'Adds to array
        DoEvents
    Loop

    'If sort then sort
    If chkSorted.Value = 1 Then QuickSort tmpAray(), LBound(tmpAray), UBound(tmpAray) 'Sort the data
    
    Select Case cboOutput.ListIndex
        Dim strFileName As String
        
        'Case 0
        Case 1
            GetSaveName hwnd, "Save", strFileName
    
            'Writes the string to a file
            If strFileName <> "" Then 'If strFileName is not empty
                Open strFileName For Output As #1
                    For TotalLinks = 1 To TotalLinks
                        Print #1, tmpAray(TotalLinks)
                    Next TotalLinks
                Close #1
            End If
            
            'Shell out to default text editor for viewing
            'apiError = ShellExecute(hWnd, "open", Dirs.AppPath & "\ie5history.txt", vbNullString, vbNullString, SW_SHOWNORMAL)
            'If apiError <= 32 Then
            '    Errors.Errors apiError, "ShellExecute"
            'End If
        Case 2
            
            GetSaveName hwnd, "Save", strFileName
    
            'Writes the string to a file
            If strFileName <> "" Then 'If strFileName is not empty
                Open strFileName For Output As #1
                    'Adds leading info
                    Print #1, "<!DOCTYPE HTML PUBLIC " & Chr$(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr$(34) & ">"
                    Print #1, "<html><head><title>IE History</title></head><body>"
                    Print #1, "<meta http-equiv=" & Chr$(34) & "Content-Type" & Chr$(34) & " content=" & Chr$(34) & "text/html; charset=iso-8859-1" & Chr$(34) & ">"
                    Print #1, "<body bgcolor=" & Chr$(34) & "#000000" & Chr$(34) & " text=" & Chr$(34) & "#D4D4D4" & Chr$(34) & " link=" & Chr$(34) & "#FFFFFF" & Chr$(34) & " vlink=" & Chr$(34) & "#D4D4D4" & Chr$(34) & " alink=" & Chr$(34) & "#D4D4D4" & Chr$(34) & ">"
                    Print #1, "Total Number Of Links: " & TotalLinks & "<br><br>"
            
                    For TotalLinks = 1 To TotalLinks 'Cycles through array
                        Print #1, "<a href=" & Chr$(34) & tmpAray(TotalLinks) & Chr$(34) & ">" & tmpAray(TotalLinks) & "</a><br><br>"
                    Next TotalLinks
                    
                    'Adds trailing info
                    Print #1, "</div></body></html>"
                Close #1
            End If
    End Select
End Sub

Private Sub Form_Load()
    With cboOutput
        .AddItem "None"
        .AddItem "Text"
        .AddItem "HTML 4"
    End With
    
    cboOutput.ListIndex = GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira\IEHistory", "Output")
    chkSorted.Value = GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira\IEHistory", "Sorted")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\IEHistory", "Output", cboOutput.ListIndex
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\IEHistory", "Sorted", chkSorted.Value
    
    Unload Me
End Sub
