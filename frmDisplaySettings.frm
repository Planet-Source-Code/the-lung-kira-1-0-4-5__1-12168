VERSION 5.00
Begin VB.Form frmDisplaySettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Display Settings"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "frmDisplaySettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboModes 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CheckBox chkGlobal 
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.ComboBox cboRate 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2400
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblGlobal 
      Caption         =   "Global Change"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblAvailable 
      Caption         =   "Modes Available"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblRate 
      Caption         =   "Refresh Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmDisplaySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DEVMODE As DEVMODE

Private Sub cmdApply_Click()
    With DEVMODE
        .dmSize = Len(DEVMODE)
        .dmBitsPerPel = CLng(Right$(cboModes.List(cboModes.ListIndex), 2))
        .dmPelsWidth = CLng(Trim$(Left$(cboModes.List(cboModes.ListIndex), 8)))
        .dmPelsHeight = CLng(Trim$(Mid(cboModes.List(cboModes.ListIndex), 8, 8)))
        .dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_DISPLAYFREQUENCY
        .dmDisplayFrequency = CLng(cboRate.Text)
        '.dmPosition 'Multimonitor
    End With
    
    'Test
    If ChangeDisplaySettings(DEVMODE, CDS_TEST) <> 0 Then
        MsgBox "Display test failed. Mode was not set.", vbExclamation, "Error"
        Exit Sub 'If error exit here
    End If
    
    apiError = chkGlobal.Value
    If apiError = 1 Then
        apiError = ChangeDisplaySettings(DEVMODE, CDS_UPDATEREGISTRY Or CDS_GLOBAL)
    Else
        apiError = ChangeDisplaySettings(DEVMODE, CDS_UPDATEREGISTRY)
    End If
    Select Case apiError 'What to do
        Case DISP_CHANGE_RESTART:    MsgBox "Must restart computer for changes to be implemented.", vbInformation, "Restart"
        Case DISP_CHANGE_BADFLAGS:   Failed "ChangeDisplaySettings"
        Case DISP_CHANGE_BADPARAM:   Failed "ChangeDisplaySettings"
        Case DISP_CHANGE_FAILED:     Failed "ChangeDisplaySettings"
        Case DISP_CHANGE_BADMODE:    Failed "ChangeDisplaySettings"
        Case DISP_CHANGE_NOTUPDATED: Failed "ChangeDisplaySettings"
    End Select
    
    'Change screenedges accordingly
    ScreenEdge.X = Screen.Width \ Screen.TwipsPerPixelX
    ScreenEdge.Y = Screen.Height \ Screen.TwipsPerPixelY
End Sub

Private Sub Form_Load()
    Dim lngIncrement As Long
    
    Dim curBPP As Integer
    Dim curWidth As Integer
    Dim curHeight As Integer

    DEVMODE.dmSize = Len(DEVMODE)
    
    'Get current settings
    curBPP = GetDeviceCaps(hdc, BITSPIXEL)
    curWidth = Screen.Width \ Screen.TwipsPerPixelX
    curHeight = Screen.Height \ Screen.TwipsPerPixelY

    Do 'Get all possible display modes
        If EnumDisplaySettings(ByVal 0, lngIncrement, DEVMODE) = 0 Then
            Failed "EnumDisplaySettings"
            Exit Do
        End If
        
        lngIncrement = lngIncrement + 1 'Increment
        
        cboModes.AddItem Left$(DEVMODE.dmPelsWidth & Space$(8), 8) & Left$(DEVMODE.dmPelsHeight & Space$(8), 8) & DEVMODE.dmBitsPerPel
        
        If curBPP = DEVMODE.dmBitsPerPel And _
            curWidth = DEVMODE.dmPelsWidth And _
            curHeight = DEVMODE.dmPelsHeight Then
            cboModes.ListIndex = cboModes.NewIndex
        End If
    Loop
    
    For lngIncrement = 1 To 300
        cboRate.AddItem lngIncrement
    Next lngIncrement
    
    If WinID = "WIN32_NT" Then
        Dim curVRefresh As Integer
        curVRefresh = GetDeviceCaps(hdc, VREFRESH)
        
        If curVRefresh > 1 Then '0 or 1 = default
            cboRate.ListIndex = curVRefresh - 1
        End If
    Else
        Dim strData As String
        strData = GetSettingString(HKEY_CURRENT_CONFIG, "Display\Settings", "RefreshRate")
        
        If strData <> "" Then cboRate.ListIndex = CLng(strData) - 1
    End If
    
    If cboRate.ListIndex = 0 Then
        cboRate.ListIndex = 60 'Default
    End If
    
    chkGlobal.Value = GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira\DisplaySettings", "GlobalChange")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\DisplaySettings", "GlobalChange", chkGlobal.Value
    Unload Me
End Sub
