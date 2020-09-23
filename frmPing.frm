VERSION 5.00
Begin VB.Form frmPing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ping"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmPing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstRoundTripTime 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1935
   End
   Begin VB.HScrollBar hsTTL 
      Height          =   135
      LargeChange     =   5
      Left            =   2760
      Max             =   255
      TabIndex        =   10
      Top             =   1155
      Value           =   128
      Width           =   1335
   End
   Begin VB.HScrollBar hsTimeout 
      Height          =   135
      LargeChange     =   5
      Left            =   1440
      TabIndex        =   7
      Top             =   1150
      Value           =   5000
      Width           =   1215
   End
   Begin VB.HScrollBar hsNumber 
      Height          =   135
      LargeChange     =   5
      Left            =   120
      TabIndex        =   4
      Top             =   1150
      Value           =   1
      Width           =   1215
   End
   Begin VB.TextBox txtAvg 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtFailed 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtTTL 
      Height          =   285
      Left            =   2760
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtTimeout 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "Ping"
      Height          =   350
      Left            =   3120
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblAvg 
      Caption         =   "Avg"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblFailed 
      Caption         =   "Failed"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblTTL 
      Caption         =   "Time To Live"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblNumber 
      Caption         =   "Number of Pings"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblRoundTripTime 
      Caption         =   "Round Trip Time"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblIP 
      Caption         =   "IP"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblTimeout 
      Caption         =   "Timeout"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPing_Click()
    'Clear
    lstRoundTripTime.Clear
    txtAvg.Text = ""
    txtFailed.Text = ""
    
    Dim hFile As Long
    Dim strData As String
    Dim lngIncrement As Long
    
    Dim Avg() As Long
    Dim numAvg As Long
    Dim Failed As Long
    
    Dim IP_OPTION_INFORMATION As IP_OPTION_INFORMATION
    Dim ICMP_ECHO_REPLY As ICMP_ECHO_REPLY
    
    hFile = IcmpCreateFile() 'Create icmp handle
    IP_OPTION_INFORMATION.TTL = hsTTL.Value
    'IP_OPTION_INFORMATION.Tos = 8
    strData = "ICMP ECHO DATA"
    
    For lngIncrement = 1 To hsNumber.Value 'Cycle through pings
        apiError = IcmpSendEcho(hFile, inet_addr(txtIP.Text & Chr$(0)), strData, Len(strData), IP_OPTION_INFORMATION, ICMP_ECHO_REPLY, Len(ICMP_ECHO_REPLY), hsTimeout.Value)
        If apiError = 0 Then
            lstRoundTripTime.AddItem Left$(lngIncrement & Space$(7), 7) & "Failed"
            Failed = Failed + 1 'Increment
        Else
            numAvg = numAvg + 1 'Increment
            ReDim Preserve Avg(numAvg) 'Resizes array without destroying
            
            'Dump info back
            Avg(numAvg) = ICMP_ECHO_REPLY.RoundTripTime
            lstRoundTripTime.AddItem Left$(lngIncrement & Space$(7), 7) & ICMP_ECHO_REPLY.RoundTripTime
        End If
        
        DoEvents
    Next lngIncrement
    
    Dim dblAvg As Double
    For lngIncrement = 1 To numAvg 'Cycle array
        dblAvg = dblAvg + Avg(numAvg) 'Dump entire array to double
    Next lngIncrement
    If dblAvg > 0 Then dblAvg = dblAvg / numAvg
    
    txtAvg.Text = dblAvg
    txtFailed.Text = Round(((Failed / hsNumber.Value) * 100), 0) & "%"
    
    Call IcmpCloseHandle(hFile) 'Close icmp handle
End Sub

Private Sub Form_Load()
    hsNumber.Value = GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira\Ping", "Number")
    hsTimeout.Value = GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira\Ping", "Timeout")
    hsTTL.Value = GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira\Ping", "TTL")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\Ping", "Number", hsNumber.Value
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\Ping", "Timeout", hsTimeout.Value
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\Ping", "TTL", hsTTL.Value
    
    Unload Me
End Sub

Private Sub hsNumber_Change()
    txtNumber.Text = hsNumber.Value
End Sub

Private Sub hsTimeout_Change()
    txtTimeout.Text = hsTimeout.Value
End Sub

Private Sub hsTTL_Change()
    txtTTL.Text = hsTTL.Value
End Sub

Private Sub txtNumber_Change()
    On Error Resume Next
    
    If CInt(txtNumber.Text) <= 0 Then txtNumber.Text = "1" 'If less than 0 resets to min , also does error trapping
    hsNumber.Value = CInt(txtNumber.Text)
End Sub

Private Sub txtTimeout_Change()
    On Error Resume Next
    
    If CInt(txtTimeout.Text) < 0 Then txtTimeout.Text = "0" 'If less than 0 resets to min , also does error trapping
    hsTimeout.Value = CInt(txtTimeout.Text)
End Sub

Private Sub txtTTL_Change()
    On Error Resume Next
    
    If CByte(txtTTL.Text) < 0 Then txtTTL.Text = "0" 'If less than 0 resets to min , also does error trapping
    hsTTL.Value = CByte(txtTTL.Text) 'Allows custom value to be set , by converting box to int sending it to slider
End Sub
