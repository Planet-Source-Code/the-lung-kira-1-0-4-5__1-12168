VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kira"
   ClientHeight    =   360
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   1095
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   1095
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picHS 
      Height          =   255
      Left            =   360
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picHM 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu mnuMain 
      Caption         =   ""
      Begin VB.Menu mnuAccessibility 
         Caption         =   "Accessibility"
         Begin VB.Menu mnuStickyKeys 
            Caption         =   "Sticky Keys"
         End
      End
      Begin VB.Menu mnuFileFormat 
         Caption         =   "File Formats"
         Begin VB.Menu mnuMpeg 
            Caption         =   "Mpeg"
         End
         Begin VB.Menu mnuMZ 
            Caption         =   "MZ"
         End
         Begin VB.Menu mnuNE 
            Caption         =   "NE"
            Begin VB.Menu mnuSFI_NE 
               Caption         =   "String File Info Editor NE"
            End
         End
         Begin VB.Menu mnuPE 
            Caption         =   "PE"
            Begin VB.Menu mnuSFI_PE 
               Caption         =   "String File Info Editor PE"
            End
         End
      End
      Begin VB.Menu mnuHardware 
         Caption         =   "Hardware"
         Begin VB.Menu mnuCmos 
            Caption         =   "Cmos Contents"
         End
         Begin VB.Menu mnuDrives 
            Caption         =   "Drives"
            Begin VB.Menu mnuDiskSpace 
               Caption         =   "Disk Space"
            End
            Begin VB.Menu mnuVolumeInfo 
               Caption         =   "Volume Info"
            End
         End
         Begin VB.Menu mnuMemoryStatus 
            Caption         =   "Memory Status"
         End
         Begin VB.Menu mnuPowerStatus 
            Caption         =   "Power Status"
         End
         Begin VB.Menu mnuProcessor 
            Caption         =   "Processor"
            Begin VB.Menu mnuCPUID 
               Caption         =   "CPUID"
            End
            Begin VB.Menu mnuProcessorInfo 
               Caption         =   "Processor Info"
            End
         End
      End
      Begin VB.Menu mnuInternetNetwork 
         Caption         =   "Internet / Network"
         Begin VB.Menu mnuDayTime 
            Caption         =   "DayTime"
         End
         Begin VB.Menu mnuDiscard 
            Caption         =   "Discard"
         End
         Begin VB.Menu mnuGetIPHost 
            Caption         =   "Get IP/Host"
         End
         Begin VB.Menu mnuPing 
            Caption         =   "Ping"
         End
         Begin VB.Menu mnuTime 
            Caption         =   "Time"
         End
         Begin VB.Menu mnuBreak0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIP_Stats 
            Caption         =   "IP Stats"
         End
         Begin VB.Menu mnuICPM_Stats 
            Caption         =   "ICMP Stats"
         End
         Begin VB.Menu mnuNetworkInfo 
            Caption         =   "Network Info"
         End
         Begin VB.Menu mnuTCP_Stats 
            Caption         =   "TCP Stats"
         End
         Begin VB.Menu mnuUDP_Stats 
            Caption         =   "UDP Stats"
         End
         Begin VB.Menu mnuWinsockInfo 
            Caption         =   "Winsock Info"
         End
      End
      Begin VB.Menu mnuPeripherials 
         Caption         =   "Peripherials"
         Begin VB.Menu mnuKeyboard 
            Caption         =   "Keyboard"
            Begin VB.Menu mnuKeyboardSettings 
               Caption         =   "Keyboard Settings"
            End
            Begin VB.Menu mnuKeyboardInfo 
               Caption         =   "Keyboard Info"
            End
         End
         Begin VB.Menu mnuDisplay 
            Caption         =   "Display"
            Begin VB.Menu mnuDisplaySettings 
               Caption         =   "Display Settings"
            End
            Begin VB.Menu mnuMonitorInfo 
               Caption         =   "Monitor Info"
            End
         End
         Begin VB.Menu mnuMouse 
            Caption         =   "Mouse"
            Begin VB.Menu mnuMouseInfo 
               Caption         =   "Mouse Info"
            End
            Begin VB.Menu mnuMouseMov 
               Caption         =   "Mouse Movements"
            End
            Begin VB.Menu mnuMouseSettings 
               Caption         =   "Mouse Settings"
            End
            Begin VB.Menu mnuMouseWarp 
               Caption         =   "Mouse Warp"
            End
         End
      End
      Begin VB.Menu mnuSoftware 
         Caption         =   "Software"
         Begin VB.Menu mnuCDPlayer 
            Caption         =   "CD Player"
         End
         Begin VB.Menu mnuChecksum 
            Caption         =   "Checksum"
         End
         Begin VB.Menu mnuIE5 
            Caption         =   "Internet Explorer 5"
            Begin VB.Menu mnuIECache 
               Caption         =   "Cache Hit/Miss"
            End
            Begin VB.Menu mnuIEHistory 
               Caption         =   "History Viewer"
            End
            Begin VB.Menu mnuIEOptions 
               Caption         =   "IE Options"
            End
         End
      End
      Begin VB.Menu mnuWin 
         Caption         =   "Windows"
         Begin VB.Menu mnuCachedPasswords 
            Caption         =   "Cached Passwords"
         End
         Begin VB.Menu mnuDirectories 
            Caption         =   "Directories"
         End
         Begin VB.Menu mnuErrors 
            Caption         =   "Errors"
         End
         Begin VB.Menu mnuFileS 
            Caption         =   "Files"
            Begin VB.Menu mnuFileAttributes 
               Caption         =   "File Attributes"
            End
            Begin VB.Menu mnuFileTime 
               Caption         =   "File Time"
            End
            Begin VB.Menu mnuSharedFiles 
               Caption         =   "Shared Files"
            End
         End
         Begin VB.Menu mnuIcons 
            Caption         =   "Icons"
            Begin VB.Menu mnuIconInfo 
               Caption         =   "Icon Info"
            End
            Begin VB.Menu mnuIconSettings 
               Caption         =   "Icon Settings"
            End
         End
         Begin VB.Menu mnuLoaded 
            Caption         =   "Loaded"
            Begin VB.Menu mnuHeaps 
               Caption         =   "Heaps"
            End
            Begin VB.Menu mnuModules 
               Caption         =   "Modules"
            End
            Begin VB.Menu mnuProcesses 
               Caption         =   "Processes"
            End
            Begin VB.Menu mnuThreads 
               Caption         =   "Threads"
            End
         End
         Begin VB.Menu mnuMenuSettings 
            Caption         =   "Menu Settings"
         End
         Begin VB.Menu mnuPerfMon 
            Caption         =   "Performance Monitor"
         End
         Begin VB.Menu mnuStartMenu 
            Caption         =   "Start Menu"
         End
         Begin VB.Menu mnuUpTime 
            Caption         =   "Up Time"
         End
         Begin VB.Menu mnuUser 
            Caption         =   "User"
            Begin VB.Menu mnuCompUserName 
               Caption         =   "Computer/User Name"
            End
            Begin VB.Menu mnuRegistered 
               Caption         =   "Registered To"
            End
         End
         Begin VB.Menu mnuWindowInfo 
            Caption         =   "Window Info"
         End
         Begin VB.Menu mnuWindowKiller 
            Caption         =   "Window Killer"
         End
         Begin VB.Menu mnuWindows 
            Caption         =   "Windows"
         End
         Begin VB.Menu mnuWinFileProtection 
            Caption         =   "Windows File Protection"
         End
         Begin VB.Menu mnuWinInfo 
            Caption         =   "Windows Info"
         End
      End
      Begin VB.Menu mnuBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnOff 
         Caption         =   "On / Off"
         Begin VB.Menu mnuMouseMovOO 
            Caption         =   "Mouse Movements"
         End
         Begin VB.Menu mnuMouseWarpOO 
            Caption         =   "Mouse Warp"
         End
         Begin VB.Menu mnuWindowKillerOO 
            Caption         =   "Window Killer"
         End
      End
      Begin VB.Menu mnuBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Begin VB.Menu mnuAbout 
            Caption         =   "About"
         End
         Begin VB.Menu mnuAboutShell 
            Caption         =   "About Shell"
         End
      End
      Begin VB.Menu mnuBreak3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenAll 
         Caption         =   "Open All"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close All"
      End
      Begin VB.Menu mnuBreak4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitWindows 
         Caption         =   "Exit Windows"
      End
      Begin VB.Menu mnuBreak5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call App_Startup
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim wParam As Single
    wParam = X / Screen.TwipsPerPixelX
    
    Select Case wParam 'For system tray icon
        Case WM_LBUTTONUP
            apiError = SetForegroundWindow(Me.hwnd) 'Make sure its on top
            frmMain.PopupMenu mnuMain
        Case WM_RBUTTONUP
            apiError = SetForegroundWindow(Me.hwnd) 'Make sure its on top
            frmMain.PopupMenu mnuMain
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call App_Shutdown
End Sub

Private Sub mnuAbout_Click()
    frmMainAbout.Show
End Sub

Private Sub mnuAboutShell_Click()
    If ShellAbout(0&, "", "", Me.Icon) = 0 Then
        Failed "ShellAbout"
    End If
End Sub

Private Sub mnuCachedPasswords_Click()
    frmCachedPasswords.Show
End Sub

Private Sub mnuCDPlayer_Click()
    frmCDPlayer.Show
End Sub

Private Sub mnuChecksum_Click()
    frmChecksum.Show
End Sub

Private Sub mnuCloseAll_Click()
    Unload frmCachedPasswords
    DoEvents
    Unload frmCDPlayer
    DoEvents
    Unload frmChecksum
    DoEvents
    Unload frmCmos
    DoEvents
    Unload frmCompUser
    DoEvents
    Unload frmCPUID
    DoEvents
    Unload frmDayTime
    DoEvents
    Unload frmDirectories
    DoEvents
    Unload frmDiscard
    DoEvents
    Unload frmDiskSpace
    DoEvents
    Unload frmDiskVolume
    DoEvents
    Unload frmDisplaySettings
    DoEvents
    Unload frmErrors
    DoEvents
    Unload frmExitWindows
    DoEvents
    Unload frmFileAttributes
    DoEvents
    Unload frmFileTime
    DoEvents
    Unload frmGetIPHost
    DoEvents
    Unload frmHeaps
    DoEvents
    Unload frmICMP_Stats
    DoEvents
    Unload frmIconInfo
    DoEvents
    Unload frmIconSettings
    DoEvents
    Unload frmIECache
    DoEvents
    Unload frmIEHistory
    DoEvents
    Unload frmIP_Stats
    DoEvents
    Unload frmKeyboardInfo
    DoEvents
    Unload frmKeyboardSettings
    DoEvents
    Unload frmMainAbout
    DoEvents
    Unload frmMemoryStatus
    DoEvents
    Unload frmMenuSettings
    DoEvents
    Unload frmModules
    DoEvents
    Unload frmMonitorInfo
    DoEvents
    Unload frmMouseInfo
    DoEvents
    Unload frmMouseMovement
    DoEvents
    Unload frmMouseWarp
    DoEvents
    Unload frmMpeg
    DoEvents
    Unload frmMZ
    DoEvents
    Unload frmNetworkInfo
    DoEvents
    Unload frmPerfMon
    DoEvents
    Unload frmPing
    DoEvents
    Unload frmPowerStatus
    DoEvents
    Unload frmProcesses
    DoEvents
    Unload frmProcessorInfo
    DoEvents
    Unload frmRegistered
    DoEvents
    Unload frmSFI_NE
    DoEvents
    Unload frmSFI_PE
    DoEvents
    Unload frmSharedFiles
    DoEvents
    Unload frmStartMenu
    DoEvents
    Unload frmStickyKeys
    DoEvents
    Unload frmTCP_Stats
    DoEvents
    Unload frmThreads
    DoEvents
    Unload frmTime
    DoEvents
    Unload frmUDP_Stats
    DoEvents
    Unload frmUpTime
    DoEvents
    Unload frmWindowInfo
    DoEvents
    Unload frmWindowKiller
    DoEvents
    Unload frmWindows
    DoEvents
    Unload frmWinFileProtection
    DoEvents
    Unload frmWinInfo
    DoEvents
    Unload frmWinsockInfo
End Sub

Private Sub mnuCmos_Click()
    frmCmos.Show
End Sub

Private Sub mnuCompUserName_Click()
    frmCompUser.Show
End Sub

Private Sub mnuCPUID_Click()
    frmCPUID.Show
End Sub

Private Sub mnuDayTime_Click()
    frmDayTime.Show
End Sub

Private Sub mnuDirectories_Click()
    frmDirectories.Show
End Sub

Private Sub mnuDiscard_Click()
    frmDiscard.Show
End Sub

Private Sub mnuDiskSpace_Click()
    frmDiskSpace.Show
End Sub

Private Sub mnuDisplaySettings_Click()
    frmDisplaySettings.Show
End Sub

Private Sub mnuErrors_Click()
    frmErrors.Show
End Sub

Private Sub mnuExit_Click()
    Call App_Shutdown
End Sub

Private Sub mnuExitWindows_Click()
    frmExitWindows.Show
End Sub

Private Sub mnuFileAttributes_Click()
    frmFileAttributes.Show
End Sub

Private Sub mnuFileTime_Click()
    frmFileTime.Show
End Sub

Private Sub mnuHeaps_Click()
    frmHeaps.Show
End Sub

Private Sub mnuIconInfo_Click()
    frmIconInfo.Show
End Sub

Private Sub mnuIconSettings_Click()
    frmIconSettings.Show
End Sub

Private Sub mnuICPM_Stats_Click()
    frmICMP_Stats.Show
End Sub

Private Sub mnuIECache_Click()
    frmIECache.Show
End Sub

Private Sub mnuIEHistory_Click()
    frmIEHistory.Show
End Sub

Private Sub mnuIEOptions_Click()
    frmIEOptions.Show
End Sub

Private Sub mnuIP_Stats_Click()
    frmIP_Stats.Show
End Sub

Private Sub mnuKeyboardInfo_Click()
    frmKeyboardInfo.Show
End Sub

Private Sub mnuGetIPHost_Click()
    frmGetIPHost.Show
End Sub

Private Sub mnuKeyboardSettings_Click()
    frmKeyboardSettings.Show
End Sub

Private Sub mnuMemoryStatus_Click()
    frmMemoryStatus.Show
End Sub

Private Sub mnuMenuSettings_Click()
    frmMenuSettings.Show
End Sub

Private Sub mnuModules_Click()
    frmModules.Show
End Sub

Private Sub mnuMonitorInfo_Click()
    frmMonitorInfo.Show
End Sub

Private Sub mnuMouseInfo_Click()
    frmMouseInfo.Show
End Sub

Private Sub mnuMouseMov_Click()
    frmMouseMovement.Show
End Sub

Private Sub mnuMouseMovOO_Click()
    'If checked then uncheck, vice versa
    If mnuMouseMovOO.Checked = False Then 'Off to on
        mnuMouseMovOO.Checked = True
        
        Dim POINTAPI As POINTAPI
        GetCursorPos POINTAPI 'Dumps info to pointapi

        'Gives point of reference for starting
        MouseMovTmpX = POINTAPI.X
        MouseMovTmpY = POINTAPI.Y
    Else 'On to off
        mnuMouseMovOO.Checked = False
    End If
End Sub

Private Sub mnuMouseSettings_Click()
    frmMouseSettings.Show
End Sub

Private Sub mnuMouseWarp_Click()
    frmMouseWarp.Show
End Sub

Private Sub mnuMouseWarpOO_Click()
    'If checked then uncheck, vice versa
    If mnuMouseWarpOO.Checked = False Then 'Off to on
        mnuMouseWarpOO.Checked = True
        
        'Converts the twips to pixels and sets the edges
        ScreenEdge.X = Screen.Width \ Screen.TwipsPerPixelX
        ScreenEdge.Y = Screen.Height \ Screen.TwipsPerPixelY
    Else 'On to off
        mnuMouseWarpOO.Checked = False
    End If
End Sub

Private Sub mnuMpeg_Click()
    frmMpeg.Show
End Sub

Private Sub mnuMZ_Click()
    frmMZ.Show
End Sub

Private Sub mnuNetworkInfo_Click()
    frmNetworkInfo.Show
End Sub

Private Sub mnuOpenAll_Click()
    frmCachedPasswords.Show
    DoEvents
    frmCDPlayer.Show
    DoEvents
    frmChecksum.Show
    DoEvents
    frmCmos.Show
    DoEvents
    frmCompUser.Show
    DoEvents
    frmCPUID.Show
    DoEvents
    frmDayTime.Show
    DoEvents
    frmDirectories.Show
    DoEvents
    frmDiscard.Show
    DoEvents
    frmDiskSpace.Show
    DoEvents
    frmDiskVolume.Show
    DoEvents
    frmDisplaySettings.Show
    DoEvents
    frmErrors.Show
    DoEvents
    frmExitWindows.Show
    DoEvents
    frmFileAttributes.Show
    DoEvents
    frmFileTime.Show
    DoEvents
    frmGetIPHost.Show
    DoEvents
    frmHeaps.Show
    DoEvents
    frmICMP_Stats.Show
    DoEvents
    frmIconInfo.Show
    DoEvents
    frmIconSettings.Show
    DoEvents
    frmIECache.Show
    DoEvents
    frmIEHistory.Show
    DoEvents
    frmIP_Stats.Show
    DoEvents
    frmKeyboardInfo.Show
    DoEvents
    frmKeyboardSettings.Show
    DoEvents
    frmMainAbout.Show
    DoEvents
    frmMemoryStatus.Show
    DoEvents
    frmMenuSettings.Show
    DoEvents
    frmModules.Show
    DoEvents
    frmMonitorInfo.Show
    DoEvents
    frmMouseInfo.Show
    DoEvents
    frmMouseMovement.Show
    DoEvents
    frmMouseWarp.Show
    DoEvents
    frmMpeg.Show
    DoEvents
    frmMZ.Show
    DoEvents
    frmNetworkInfo.Show
    DoEvents
    frmPerfMon.Show
    DoEvents
    frmPing.Show
    DoEvents
    frmPowerStatus.Show
    DoEvents
    frmProcesses.Show
    DoEvents
    frmProcessorInfo.Show
    DoEvents
    frmRegistered.Show
    DoEvents
    frmSFI_NE.Show
    DoEvents
    frmSFI_PE.Show
    DoEvents
    frmSharedFiles.Show
    DoEvents
    frmStartMenu.Show
    DoEvents
    frmStickyKeys.Show
    DoEvents
    frmTCP_Stats.Show
    DoEvents
    frmThreads.Show
    DoEvents
    frmTime.Show
    DoEvents
    frmUDP_Stats.Show
    DoEvents
    frmUpTime.Show
    DoEvents
    frmWindowInfo.Show
    DoEvents
    frmWindowKiller.Show
    DoEvents
    frmWindows.Show
    DoEvents
    frmWinFileProtection.Show
    DoEvents
    frmWinInfo.Show
    DoEvents
    frmWinsockInfo.Show
End Sub

Private Sub mnuPerfMon_Click()
    frmPerfMon.Show
End Sub

Private Sub mnuPing_Click()
    frmPing.Show
End Sub

Private Sub mnuPowerStatus_Click()
    frmPowerStatus.Show
End Sub

Private Sub mnuProcesses_Click()
    frmProcesses.Show
End Sub

Private Sub mnuProcessorInfo_Click()
    frmProcessorInfo.Show
End Sub

Private Sub mnuRegistered_Click()
    frmRegistered.Show
End Sub

Private Sub mnuSFI_NE_Click()
    frmSFI_NE.Show
End Sub

Private Sub mnuSFI_PE_Click()
    frmSFI_PE.Show
End Sub

Private Sub mnuSharedFiles_Click()
    frmSharedFiles.Show
End Sub

Private Sub mnuStartMenu_Click()
    frmStartMenu.Show
End Sub

Private Sub mnuStickyKeys_Click()
    frmStickyKeys.Show
End Sub

Private Sub mnuTCP_Stats_Click()
    frmTCP_Stats.Show
End Sub

Private Sub mnuThreads_Click()
    frmThreads.Show
End Sub

Private Sub mnuTime_Click()
    frmTime.Show
End Sub

Private Sub mnuUDP_Stats_Click()
    frmUDP_Stats.Show
End Sub

Private Sub mnuUpTime_Click()
    frmUpTime.Show
End Sub

Private Sub mnuVolumeInfo_Click()
    frmDiskVolume.Show
End Sub

Private Sub mnuWindowInfo_Click()
    frmWindowInfo.Show
End Sub

Private Sub mnuWindowKiller_Click()
    frmWindowKiller.Show
End Sub

Private Sub mnuWindowKillerOO_Click()
    'If checked then uncheck, vice versa
    If mnuWindowKillerOO.Checked = False Then 'Off to on
        mnuWindowKillerOO.Checked = True
    Else 'On to off
        mnuWindowKillerOO.Checked = False
    End If
End Sub

Private Sub mnuWindows_Click()
    frmWindows.Show
End Sub

Private Sub mnuWinFileProtection_Click()
    frmWinFileProtection.Show
End Sub

Private Sub mnuWinInfo_Click()
    frmWinInfo.Show
End Sub

Private Sub mnuWinsockInfo_Click()
    frmWinsockInfo.Show
End Sub
