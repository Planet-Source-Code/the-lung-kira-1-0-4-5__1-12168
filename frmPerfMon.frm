VERSION 5.00
Begin VB.Form frmPerfMon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Performance Monitor"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmPerfMon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.Timer timerPerfMon 
      Enabled         =   0   'False
      Left            =   2760
      Top             =   720
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox txtInterval 
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.HScrollBar hsInterval 
      Height          =   255
      LargeChange     =   5
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Value           =   1000
      Width           =   2535
   End
   Begin VB.ComboBox cboObject 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblData 
      Caption         =   "Data"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblInterval 
      Caption         =   "Update Interval"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblObject 
      Caption         =   "Object"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmPerfMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'9x
Dim Differentiate As Boolean
Dim oldData As Long

'NT
Private Type CounterInfo
    hCounter As Long
    strName As String
End Type

Dim Counters() As CounterInfo
Dim numCounters As Long
Dim hQuery As Long

Private Sub cboObject_Click()
    If WinID = "WIN32_WINDOWS" Then '9x
        Dim extPath As String
        Dim extPath2 As String
        Dim pos As Integer
        
        Dim tmpString As String
        tmpString = cboObject.List(cboObject.ListIndex)
        
        pos = InStr(1, tmpString, "\")
        extPath = Left$(tmpString, pos) 'Left side
        extPath2 = Right$(tmpString, Len(tmpString) - pos) 'Right side
        
        'Differentiate value each cycle 0/1
        If GetSettingString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\PerfStats\Enum\" & extPath & extPath2, "Differentiate") = "TRUE" Then
            Differentiate = True
        Else
            Differentiate = False
        End If
        
        'Name / Description
        txtName.Text = GetSettingString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\PerfStats\Enum\" & extPath & extPath2, "Name")
        txtDescription.Text = GetSettingString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\PerfStats\Enum\" & extPath & extPath2, "Description")
        
        oldData = CLng(GetSettingBinary(HKEY_DYN_DATA, "PerfStats\StatData", cboObject.List(cboObject.ListIndex)))
    'Else 'NT
    
    End If

    txtData.Text = ""
    
    timerPerfMon.Interval = hsInterval.Value
    timerPerfMon.Enabled = True
    timerPerfMon_Timer
End Sub

Private Sub cmdAdd_Click()
    Dim retLen As Long
    Dim strCounterPath As String * 256
    Dim strCounterName As String
    Dim hCounter As Long
    
    retLen = PdhVbGetOneCounterPath(strCounterPath, 256, PERF_DETAIL_WIZARD, "Add Object")
    If retLen = 0 Then Exit Sub
    
    strCounterName = Left$(strCounterPath, retLen)
    
    apiError = PdhVbAddCounter(hQuery, strCounterName, hCounter)
    If apiError <> 0 Then pdhmsg.PdhError apiError, "PdhVbAddCounter"
    
    ReDim Preserve Counters(numCounters)
    Counters(numCounters).hCounter = hCounter
    Counters(numCounters).strName = strCounterName
    numCounters = numCounters + 1 'Increment
    
    cboObject.AddItem strCounterName
End Sub

Private Sub Form_Load()
    With cboObject
        .Clear
        
        If WinID = "WIN32_WINDOWS" Then '9x
            Dim lngCount As Long
            Dim strValueName() As String
            Dim lngValueType As Long
            Dim lngIncrement As Long
            
            'Enumerate all objects
            EnumValue HKEY_DYN_DATA, "PerfStats\StatData", strValueName(), lngCount, lngValueType
            
            For lngIncrement = 0 To lngCount - 2 'Cycle through
                .AddItem strValueName(lngIncrement) 'Dump
            Next lngIncrement
            
            
            Dim srvValueName() As String
            Dim srvCount As Long
            EnumValue HKEY_DYN_DATA, "PerfStats\StartSrv", srvValueName(), srvCount, lngValueType
            
            For lngIncrement = 0 To srvCount - 2
                'Start up monitoring services
                GetSettingBinary HKEY_DYN_DATA, "PerfStats\StartSrv", srvValueName(lngIncrement)
            Next lngIncrement
            
            cmdAdd.Enabled = False
        Else 'NT
            'Dim bufObjects As String
            'Dim lenBuffer As Long
            
            'bufObjects = Space$(4095) & Chr$(0)
            'lenBuffer = 4096
            
            'apiError = PdhEnumObjects(0&, 0&, bufObjects, lenBuffer, PERF_DETAIL_WIZARD, True)
            'If apiError <> 0 Then pdhmsg.PdhError apiError, "PdhEnumObjects"
            'MsgBox bufObjects
            
            apiError = PdhOpenQuery(0, 1, hQuery)
            If apiError <> 0 Then pdhmsg.PdhError apiError, "PdhOpenQuery"
            
            lblName.Enabled = False
            lblDescription.Enabled = False
        End If
    End With
    
    hsInterval.Value = GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira\PerfMon", "Interval")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerPerfMon.Enabled = True 'Disable timer
    
    If WinID = "WIN32_WINDOWS" Then '9x
        Dim srvValueName() As String
        Dim srvCount As Long
        Dim lngValueType As Long
        EnumValue HKEY_DYN_DATA, "PerfStats\StopSrv", srvValueName(), srvCount, lngValueType
        
        Dim lngIncrement As Long
        For lngIncrement = 0 To srvCount - 2
            'Stop up monitoring services
            GetSettingBinary HKEY_DYN_DATA, "PerfStats\StopSrv", srvValueName(lngIncrement)
        Next lngIncrement
    Else 'NT
        apiError = PdhCloseQuery(hQuery)
        If apiError <> 0 Then pdhmsg.PdhError apiError, "PdhCloseQuery"
    End If
    
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\PerfMon", "Interval", hsInterval.Value
    Unload Me
End Sub

Private Sub hsInterval_Change()
    txtInterval.Text = hsInterval.Value
    timerPerfMon.Interval = hsInterval.Value
End Sub

Private Sub timerPerfMon_Timer()
    If WinID = "WIN32_WINDOWS" Then '9x
        If Differentiate = True Then
            txtData.Text = Abs(CLng(GetSettingBinary(HKEY_DYN_DATA, "PerfStats\StatData", cboObject.List(cboObject.ListIndex))) - oldData)
        Else
            txtData.Text = GetSettingBinary(HKEY_DYN_DATA, "PerfStats\StatData", cboObject.List(cboObject.ListIndex))
        End If
        
        oldData = CLng(GetSettingBinary(HKEY_DYN_DATA, "PerfStats\StatData", cboObject.List(cboObject.ListIndex)))
    Else 'NT
        apiError = PdhCollectQueryData(hQuery)
        If apiError <> 0 Then
            pdhmsg.PdhError apiError, "PdhCollectQueryData"
            timerPerfMon.Enabled = False
        End If
        
        'Return rounded off double
        txtData.Text = Round(PdhVbGetDoubleCounterValue(Counters(cboObject.ListIndex).hCounter, apiError), 0)
        If apiError <> 0 Then
            pdhmsg.PdhError apiError, "PdhVbGetDoubleCounterValue"
            timerPerfMon.Enabled = False
        End If
    End If
End Sub

Private Sub txtInterval_Change()
    On Error Resume Next
    
    If CInt(txtInterval.Text) < 0 Then txtInterval.Text = "0"
    hsInterval.Value = CInt(txtInterval.Text)
End Sub
