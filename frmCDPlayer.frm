VERSION 5.00
Begin VB.Form frmCDPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CD Player"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmCDPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFFRW 
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   2880
      Width           =   1335
   End
   Begin VB.HScrollBar hsFFRW 
      Height          =   135
      LargeChange     =   5
      Left            =   120
      Max             =   30000
      Min             =   1
      TabIndex        =   25
      Top             =   3195
      Value           =   1000
      Width           =   1335
   End
   Begin VB.Timer timerCD 
      Interval        =   950
      Left            =   2760
      Top             =   960
   End
   Begin VB.TextBox txtTotalRemaining 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtTotalElapsed 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtTrackRemaining 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtTrackElapsed 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtAlbum 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1800
      Width           =   4695
   End
   Begin VB.ComboBox cboTrack 
      Height          =   315
      Left            =   720
      TabIndex        =   22
      Top             =   2160
      Width           =   4695
   End
   Begin VB.ComboBox cboDrive 
      Height          =   315
      Left            =   720
      TabIndex        =   18
      Top             =   1440
      Width           =   4695
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   350
      Left            =   4680
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   350
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   350
      Left            =   2760
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   350
      Left            =   4680
      TabIndex        =   16
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdEject 
      Caption         =   "Eject"
      Height          =   350
      Left            =   3840
      TabIndex        =   15
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdTrackForward 
      Caption         =   "> |"
      Height          =   350
      Left            =   4920
      TabIndex        =   14
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   ">"
      Height          =   350
      Left            =   4440
      TabIndex        =   13
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      Height          =   350
      Left            =   3240
      TabIndex        =   12
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton cmdTrackBack 
      Caption         =   "| <"
      Height          =   350
      Left            =   2760
      TabIndex        =   11
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblFFRW 
      Caption         =   "FF RW Increment"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblTotalRemaining 
      Caption         =   "Total Remaining"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblTotalElapsed 
      Caption         =   "Total Elapsed"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblTrackRemaining 
      Caption         =   "Track Remaining"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblTrackElapsed 
      Caption         =   "Track Elapsed"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblAlbum 
      Caption         =   "Album"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblTrack 
      Caption         =   "Track"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblDrive 
      Caption         =   "Drive"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "frmCDPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurrentCD As String
Dim Msg As String * 255

Dim curTrack As Byte
Dim curIndex As Long
Dim lenTracks() As Long
Dim numTracks As Byte
Dim Min As Long
Dim Sec As Long

Private Sub cboDrive_Click()
    apiError = mciSendString("stop cd", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
    apiError = mciSendString("close cd", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
    
    CurrentCD = Left$(cboDrive.List(cboDrive.ListIndex), 3)
    
    apiError = mciSendString("open  " & CurrentCD & " type cdaudio alias cd wait shareable", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
    apiError = mciSendString("set cd time format tmsf", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
    
    
    UpdateTracks
    cboTrack.Clear
    
    Dim tmpLong As Long
    For tmpLong = 1 To numTracks
        'Clear
        Sec = 0
        Min = 0
    
        Sec = Sec + lenTracks(tmpLong)
        
        Do
            If Sec > 60 Then
                Sec = Sec - 60
                Min = Min + 1
            Else
                Exit Do
            End If
        Loop
                
        cboTrack.AddItem Right$("00" & tmpLong, 2) & "     " & Right$("00" & Min, 2) & ":" & Right$("00" & Sec, 2) & "     "
    Next tmpLong
    
    If cboTrack.ListCount > 1 Then 'If tracks
        cboTrack.ListIndex = 0 'Set to track 1
    End If
    
    timerCD_Timer
End Sub

Private Sub cboTrack_Click()
    'Check to see if CD is playing
    apiError = mciSendString("status cd mode", Msg, 255, 0)
    If apiError <> 0 Then mciError apiError
    
    curTrack = Val(Left$(cboTrack.List(cboTrack.ListIndex), 2))
    
    If Left$(Msg, 7) = "playing" Then
        apiError = mciSendString("play cd from " & curTrack, Msg, 255, 0)
        If apiError <> 0 Then mciError apiError
    Else
        apiError = mciSendString("seek cd to " & curTrack, Msg, 255, 0)
        If apiError <> 0 Then mciError apiError
    End If
    
    timerCD_Timer
End Sub

Private Sub cmdBack_Click()
    apiError = mciSendString("set cd time format milliseconds", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
    
    'Check to see if CD is playing
    apiError = mciSendString("status cd mode", Msg, 255, 0)
    If apiError <> 0 Then mciError apiError
    
    If Left$(Msg, 7) = "playing" Then
        apiError = mciSendString("status cd position wait", Msg, 255, 0)
        If apiError <> 0 Then mciError apiError
    
        apiError = mciSendString("play cd from " & CStr(CLng(Msg) - hsFFRW.Value), 0&, 0, 0)
        If apiError <> 0 Then mciError apiError
    Else
        apiError = mciSendString("status cd position wait", Msg, 255, 0)
        If apiError <> 0 Then mciError apiError
    
        apiError = mciSendString("seek cd to " & CStr(CLng(Msg) - hsFFRW.Value), 0&, 0, 0)
        If apiError <> 0 Then mciError apiError
    End If
    
    apiError = mciSendString("set cd time format tmsf", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
    
    timerCD_Timer
    
    If cboTrack.ListCount > 0 Then
    If curTrack < cboTrack.ListCount Then
        cboTrack.ListIndex = curTrack 'Show correct track
    End If
    End If
End Sub

Private Sub cmdClose_Click()
    apiError = mciSendString("status cd mode", Msg, 255, 0)
    If apiError <> 0 Then mciError apiError
    
    If Left$(Msg, 4) = "open" Then
       apiError = mciSendString("set cd door closed wait", Msg, 255, 0)
       If apiError <> 0 Then mciError apiError
    End If
End Sub

Private Sub cmdEject_Click()
    apiError = mciSendString("set cd door open wait", Msg, 255, 0)
    If apiError <> 0 Then mciError apiError
    
    timerCD_Timer
End Sub

Private Sub cmdForward_Click()
    apiError = mciSendString("set cd time format milliseconds", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
    
    'Check to see if CD is playing
    apiError = mciSendString("status cd mode", Msg, 255, 0)
    If apiError <> 0 Then mciError apiError
    
    If Left$(Msg, 7) = "playing" Then
        apiError = mciSendString("status cd position wait", Msg, 255, 0)
        If apiError <> 0 Then mciError apiError
    
        apiError = mciSendString("play cd from " & CStr(CLng(Msg) + hsFFRW.Value), 0&, 0, 0)
        If apiError <> 0 Then mciError apiError
    Else
        apiError = mciSendString("status cd position wait", Msg, 255, 0)
        If apiError <> 0 Then mciError apiError
    
        apiError = mciSendString("seek cd to " & CStr(CLng(Msg) + hsFFRW.Value), 0&, 0, 0)
        If apiError <> 0 Then mciError apiError
    End If
    
    apiError = mciSendString("set cd time format tmsf", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
    
    timerCD_Timer

    If cboTrack.ListCount > 0 Then
    If curTrack < cboTrack.ListCount Then
        cboTrack.ListIndex = curTrack 'Show correct track
    End If
    End If
End Sub

Private Sub cmdPause_Click()
    apiError = mciSendString("pause cd wait", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
End Sub

Private Sub cmdPlay_Click()
    apiError = mciSendString("play cd", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
    
    timerCD_Timer
End Sub

Private Sub cmdStop_Click()
    apiError = mciSendString("stop cd wait", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
    apiError = mciSendString("seek cd to 1 wait", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
    
    timerCD_Timer
End Sub

Private Sub cmdTrackBack_Click()
    If numTracks > 1 Then 'If tracks
        'Get the current track
        apiError = mciSendString("status cd current track", Msg, 255, 0)
        If apiError <> 0 Then mciError apiError
        
        curTrack = Left$(Msg, 2)
        
        'Check to see if CD is playing
        apiError = mciSendString("status cd mode", Msg, 255, 0)
        If apiError <> 0 Then mciError apiError
        
        If Left$(Msg, 7) = "playing" Then
            If curTrack = 1 Then
                 apiError = mciSendString("play cd from " & numTracks, Msg, 255, 0)
                 If apiError <> 0 Then mciError apiError
                 
                 cboTrack.ListIndex = cboTrack.ListCount - 1
            Else
                 apiError = mciSendString("play cd from " & curTrack - 1, Msg, 255, 0)
                 If apiError <> 0 Then mciError apiError
                 
                 cboTrack.ListIndex = cboTrack.ListIndex - 1
            End If
        Else
            If curTrack = 1 Then
                 apiError = mciSendString("seek cd to " & numTracks, Msg, 255, 0)
                 If apiError <> 0 Then mciError apiError
                 
                 cboTrack.ListIndex = cboTrack.ListCount - 1
            Else
                 apiError = mciSendString("seek cd to " & curTrack - 1, Msg, 255, 0)
                 If apiError <> 0 Then mciError apiError
                 
                 cboTrack.ListIndex = cboTrack.ListIndex - 1
            End If
        End If
    End If
    
    timerCD_Timer
End Sub

Private Sub cmdTrackForward_Click()
    If numTracks > 1 Then 'If tracks
        'Get the current track
        apiError = mciSendString("status cd current track", Msg, 255, 0)
        If apiError <> 0 Then mciError apiError
        
        curTrack = Left$(Msg, 2)
        
        'Check to see if CD is playing
        apiError = mciSendString("status cd mode", Msg, 255, 0)
        If apiError <> 0 Then mciError apiError
        
        If Left$(Msg, 7) = "playing" Then
            If curTrack = numTracks Then
                 apiError = mciSendString("play cd from 1", Msg, 255, 0)
                 If apiError <> 0 Then mciError apiError
                 
                 cboTrack.ListIndex = 0
            Else
                 apiError = mciSendString("play cd from " & curTrack + 1, Msg, 255, 0)
                 If apiError <> 0 Then mciError apiError
                 
                 cboTrack.ListIndex = cboTrack.ListIndex + 1
            End If
        Else
            If curTrack = numTracks Then
                 apiError = mciSendString("seek cd to 1", Msg, 255, 0)
                 If apiError <> 0 Then mciError apiError
                 
                 cboTrack.ListIndex = 0
            Else
                 apiError = mciSendString("seek cd to " & curTrack + 1, Msg, 255, 0)
                 If apiError <> 0 Then mciError apiError
                 
                 cboTrack.ListIndex = cboTrack.ListIndex + 1
            End If
        End If
    End If
    
    timerCD_Timer 'Update
End Sub

Private Sub Form_Load()
    'Read all settings
    hsFFRW.Value = GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira\CDPlayer", "FFRW")
    
    
    Dim strBuffer As String
    
    strBuffer = Space$(255) 'Padd
    
    apiError = GetLogicalDriveStrings(Len(strBuffer), strBuffer)
    If apiError = 0 Then Failed "GetLogicalDriveStrings"
    
    strBuffer = Left$(strBuffer, apiError)
    
    Dim pos As Byte
    Dim tmpString As String
    
    Do 'Do while not pos = pos of last null
        tmpString = Mid$(strBuffer, pos + 1, 3)
        pos = InStr(pos + 1, strBuffer, Chr$(0)) 'Move to null
        
        If tmpString <> "" Then
            If GetDriveType(tmpString) = DRIVE_CDROM Then
                apiError = mciSendString("stop cd", 0&, 0, 0)
                If apiError <> 0 Then mciError apiError
                apiError = mciSendString("close cd", 0&, 0, 0)
                If apiError <> 0 Then mciError apiError
                
                CurrentCD = tmpString
                
                apiError = mciSendString("open  " & CurrentCD & " type cdaudio alias cd wait shareable", 0&, 0, 0)
                If apiError <> 0 Then mciError apiError
                apiError = mciSendString("set cd time format tmsf", 0&, 0, 0)
                If apiError <> 0 Then mciError apiError
                
                
                UpdateTracks
                
                'Clear
                Sec = 0
                Min = 0
                
                Dim tmpLong As Long
                For tmpLong = 1 To numTracks
                    Sec = Sec + lenTracks(tmpLong)
                Next tmpLong
                
                Do
                    If Sec > 60 Then
                        Sec = Sec - 60
                        Min = Min + 1
                    Else
                        Exit Do
                    End If
                Loop
                
                cboDrive.AddItem UCase$(tmpString) & "     " & Right$("00" & Min, 2) & ":" & Right$("00" & Sec, 2) & "     " 'Toss it in

                apiError = mciSendString("stop cd", 0&, 0, 0)
                If apiError <> 0 Then mciError apiError
                apiError = mciSendString("close cd", 0&, 0, 0)
                If apiError <> 0 Then mciError apiError
                apiError = mciSendString("close all", 0&, 0, 0)
                If apiError <> 0 Then mciError apiError
            End If
        End If
    Loop While pos < (Len(strBuffer) - 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Save all settings
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\CDPlayer", "FFRW", hsFFRW.Value
    
    
    timerCD.Enabled = False
    
    apiError = mciSendString("stop cd", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
    apiError = mciSendString("close cd", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
    apiError = mciSendString("close all", 0&, 0, 0)
    If apiError <> 0 Then mciError apiError
  
    Unload Me
End Sub

Private Sub hsFFRW_Change()
    txtFFRW.Text = hsFFRW.Value 'Gives value to text box once choosen
End Sub

Private Sub timerCD_Timer()
    'Dim oldCurrentCD As String
    Dim tmpLong As Long
    Dim tmpLong2 As Long
        
    'oldCurrentCD = CurrentCD
    
    'For tmpLong = 1 To cboDrive.ListCount
    '    CurrentCD = Left$(cboDrive.List(tmpLong - 1), 3)
    '
    '    apiError = mciSendString("open  " & CurrentCD & " type cdaudio alias cd wait shareable", 0&, 0, 0)
    '    If apiError <> 0 Then mciError apiError
    '    apiError = mciSendString("set cd time format tmsf", 0&, 0, 0)
    '    If apiError <> 0 Then mciError apiError
    '
    '    UpdateTracks
    '
    '    apiError = mciSendString("status cd media present", Msg, 255, 0)
    '    If apiError <> 0 Then mciError apiError
    '
    '    If Left$(Msg, 4) = "true" Then
    '        'Clear
    '        Sec = 0
    '        Min = 0
    '
    '        For tmpLong2 = 1 To numTracks
    '            Sec = Sec + lenTracks(tmpLong2)
    '        Next tmpLong2
    '
    '        Do
    '            If Sec > 60 Then
    '                Sec = Sec - 60
    '                Min = Min + 1
    '            Else
    '                Exit Do
    '            End If
    '        Loop
    '
    '        cboDrive.List(tmpLong - 1) = UCase$(CurrentCD) & "     " & Right$("00" & Min, 2) & ":" & Right$("00" & Sec, 2) & "     " 'Toss it in
    '
    '
    '        If CurrentCD = oldCurrentCD Then
    '            For tmpLong2 = 1 To numTracks
    '                'Clear
    '                Sec = 0
    '                Min = 0
    '
    '                Sec = Sec + lenTracks(tmpLong2)
    '
    '                Do
    '                    If Sec > 60 Then
    '                        Sec = Sec - 60
    '                        Min = Min + 1
    '                    Else
    '                        Exit Do
    '                    End If
    '                Loop
    '
    '                cboTrack.List(tmpLong2 - 1) = Right$("00" & tmpLong2, 2) & "     " & Right$("00" & Min, 2) & ":" & Right$("00" & Sec, 2) & "     "
    '            Next tmpLong2
    '        End If
    '    Else
    '        cboDrive.List(tmpLong - 1) = UCase$(CurrentCD) & "     " & "00:00" & "     " 'Toss it in
    '
    '        If CurrentCD = oldCurrentCD Then
    '            cboTrack.Clear
    '        End If
    '    End If
    '
    '    apiError = mciSendString("close cd", 0&, 0, 0)
    '    If apiError <> 0 Then mciError apiError
    'Next tmpLong
    '
    '
    'Reset to old data
    'CurrentCD = oldCurrentCD
    '
    'apiError = mciSendString("open  " & CurrentCD & " type cdaudio alias cd wait shareable", 0&, 0, 0)
    'If apiError <> 0 Then mciError apiError
    'apiError = mciSendString("set cd time format tmsf", 0&, 0, 0)
    'If apiError <> 0 Then mciError apiError
    '
    'apiError = mciSendString("status cd mode", Msg, 255, 0)
    'If apiError <> 0 Then mciError apiError
    'If Left$(Msg, 7) = "playing" Then
    '    'apiError = mciSendString("play cd from " & CStr(curIndex), 0&, 0, 0)
    '    'If apiError <> 0 Then mciError apiError
    'Else
    '    apiError = mciSendString("seek cd to " & CStr(curIndex), 0&, 0, 0)
    '    If apiError <> 0 Then mciError apiError
    'End If
    '
    'UpdateTracks
    'cboDrive.Text = cboDrive.List(cboDrive.ListIndex)
    'cboTrack.Text = cboTrack.List(cboTrack.ListIndex)
    
    
    'Check if CD is in the player
    apiError = mciSendString("status cd media present", Msg, 255, 0)
    If apiError <> 0 Then mciError apiError
    
    If Left$(Msg, 4) = "true" Then
        Dim TrackElapsed As Long
        Dim TrackRemaining As Long
        Dim TotalElapsed As Long
        
        'Get the current track
        DoEvents
        apiError = mciSendString("status cd current track", Msg, 255, 0)
        If apiError <> 0 Then mciError apiError
        
        curTrack = CByte(Left$(Msg, 2))
        
        
        apiError = mciSendString("status cd position", Msg, 255, 0)
        If apiError <> 0 Then mciError apiError
        
        txtTrackElapsed.Text = Mid$(Msg, 4, 2) & ":" & Mid$(Msg, 7, 2)
        TrackElapsed = ((Left$(txtTrackElapsed.Text, 2) * 60) + Right$(txtTrackElapsed.Text, 2))
        curIndex = TrackElapsed
        
        'Clear
        Sec = 0
        Min = 0
        
        Sec = (lenTracks(curTrack) - TrackElapsed)
        Do
            If Sec > 60 Then
                Sec = Sec - 60
                Min = Min + 1
            Else
                Exit Do
            End If
        Loop
        
        txtTrackRemaining.Text = Right$("00" & Min, 2) & ":" & Right$("00" & Sec, 2)
        TrackRemaining = ((Left$(txtTrackRemaining.Text, 2) * 60) + Right$(txtTrackRemaining.Text, 2))
        
        
        'Clear
        Sec = 0
        Min = 0
        
        For tmpLong = 1 To curTrack
            Sec = Sec + lenTracks(tmpLong)
        Next tmpLong
        
        Sec = Sec - TrackRemaining
        Do
            If Sec > 60 Then
                Sec = Sec - 60
                Min = Min + 1
            Else
                Exit Do
            End If
        Loop
        
        txtTotalElapsed.Text = Right$("00" & Min, 2) & ":" & Right$("00" & Sec, 2)
        TotalElapsed = ((Left$(txtTotalElapsed.Text, 2) * 60) + Right$(txtTotalElapsed.Text, 2))
        
        
        'Clear
        Sec = 0
        Min = 0
        
        For tmpLong = 1 To numTracks
            Sec = Sec + lenTracks(tmpLong)
        Next tmpLong
        
        Sec = Sec - TotalElapsed
        Do
            If Sec > 60 Then
                Sec = Sec - 60
                Min = Min + 1
            Else
                Exit Do
            End If
        Loop
        
        txtTotalRemaining.Text = Right$("00" & Min, 2) & ":" & Right$("00" & Sec, 2)
    End If
End Sub

Private Sub txtFFRW_Change()
    On Error Resume Next
    
    If CInt(txtFFRW.Text) < 1 Then txtFFRW.Text = "1" 'If less than 1 resets to min , also does error trapping
    If CInt(txtFFRW.Text) > 30000 Then txtFFRW.Text = "30000" 'If greater than 30000 resets to max
    
    hsFFRW.Value = CInt(txtFFRW.Text) 'Allows custom value to be set , by converting box to int sending it to slider
End Sub


Private Sub UpdateTracks()
    ReDim Preserve lenTracks(0) 'Clear
    
    
    'Number of tracks
    apiError = mciSendString("status cd number of tracks wait", Msg, 255, 0)
    If apiError <> 0 Then mciError apiError
    
    numTracks = Val(Left$(Msg, 2))
    
    Dim tmpLong As Long
    For tmpLong = 1 To numTracks
        apiError = mciSendString("status cd length track " & tmpLong, Msg, 255, 0)
        If apiError <> 0 Then mciError apiError
        
        ReDim Preserve lenTracks(tmpLong)
        lenTracks(tmpLong) = ((Val(Left$(Msg, 2)) * 60) + Val(Mid$(Msg, 4, 2)))
    Next tmpLong
End Sub
