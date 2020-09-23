Attribute VB_Name = "dllKira"
Option Explicit

'CPUID
Public Declare Sub cpu_id Lib "kira_cpu.dll" Alias "_cpu_id" ()
Public Declare Function cpuid_APICOnChip Lib "kira_cpu.dll" Alias "_cpuid_APICOnChip" () As Integer
Public Declare Function cpuid_CMOV Lib "kira_cpu.dll" Alias "_cpuid_CMOV" () As Integer
Public Declare Function cpuid_CMPXCHG8B Lib "kira_cpu.dll" Alias "_cpuid_CMPXCHG8B" () As Integer
Public Declare Function cpuid_DebuggingExtensions Lib "kira_cpu.dll" Alias "_cpuid_DebuggingExtensions" () As Integer
Public Declare Function cpuid_Family Lib "kira_cpu.dll" Alias "_cpuid_Family" () As Integer
Public Declare Function cpuid_FGPAT Lib "kira_cpu.dll" Alias "_cpuid_FGPAT" () As Integer
Public Declare Function cpuid_FpuPresent Lib "kira_cpu.dll" Alias "_cpuid_FpuPresent" () As Integer
Public Declare Function cpuid_FXSR Lib "kira_cpu.dll" Alias "_cpuid_FXSR" () As Integer
Public Declare Function cpuid_MachineCheckException Lib "kira_cpu.dll" Alias "_cpuid_MachineCheckException" () As Integer
Public Declare Function cpuid_MCA Lib "kira_cpu.dll" Alias "_cpuid_MCA" () As Integer
Public Declare Function cpuid_MMX Lib "kira_cpu.dll" Alias "_cpuid_MMX" () As Integer
Public Declare Function cpuid_Model Lib "kira_cpu.dll" Alias "_cpuid_Model" () As Integer
Public Declare Function cpuid_MSR Lib "kira_cpu.dll" Alias "_cpuid_MSR" () As Integer
Public Declare Function cpuid_MTRR Lib "kira_cpu.dll" Alias "_cpuid_MTRR" () As Integer
Public Declare Function cpuid_PageSizeExtensions Lib "kira_cpu.dll" Alias "_cpuid_PageSizeExtensions" () As Integer
Public Declare Function cpuid_PGE Lib "kira_cpu.dll" Alias "_cpuid_PGE" () As Integer
Public Declare Function cpuid_PhysicalAddressExtensions Lib "kira_cpu.dll" Alias "_cpuid_PhysicalAddressExtensions" () As Integer
Public Declare Function cpuid_PN Lib "kira_cpu.dll" Alias "_cpuid_PN" () As Integer
Public Declare Function cpuid_PSE36 Lib "kira_cpu.dll" Alias "_cpuid_PSE36" () As Integer
Public Declare Function cpuid_SEP Lib "kira_cpu.dll" Alias "_cpuid_SEP" () As Integer
Public Declare Function cpuid_Stepping Lib "kira_cpu.dll" Alias "_cpuid_Stepping" () As Integer
Public Declare Function cpuid_TimeStampCounter Lib "kira_cpu.dll" Alias "_cpuid_TimeStampCounter" () As Integer
Public Declare Function cpuid_Type Lib "kira_cpu.dll" Alias "_cpuid_Type" () As Integer
Public Declare Function cpuid_VME Lib "kira_cpu.dll" Alias "_cpuid_VME" () As Integer
Public Declare Function cpuid_XMM Lib "kira_cpu.dll" Alias "_cpuid_XMM" () As Integer

Public Declare Function cpuid_Reserved1 Lib "kira_cpu.dll" Alias "_cpuid_Reserved1" () As Long
Public Declare Function cpuid_reserved2 Lib "kira_cpu.dll" Alias "_cpuid_reserved2" () As Integer
Public Declare Function cpuid_reserved3 Lib "kira_cpu.dll" Alias "_cpuid_reserved3" () As Integer
Public Declare Function cpuid_reserved4 Lib "kira_cpu.dll" Alias "_cpuid_reserved4" () As Integer

Public Declare Function cycles_elapsed Lib "kira_cpu.dll" Alias "_cycles_elapsed" () As Double
Public Declare Function cpuspeed_mhz Lib "kira_cpu.dll" Alias "_cpuspeed_mhz" () As Double
Public Declare Function cpuid_avail Lib "kira_cpu.dll" Alias "_cpuid_avail" () As Boolean

Public Declare Sub MouseHookInit Lib "kira_ext.dll" (ByVal hwnd As Long, ByVal hHook As Long)

    Public dllInstance As Long
    Public oldMH_Proc As Long
    
    Public mh_ProcAddress As Long
    Public mh_Hook As Long
    Public sh_RetMsg As Long

Public Sub InstallMouseHook()
    apiError = GetProcAddress(dllInstance, "MouseHookProc")
    If apiError = 0 Then
        Failed "GetProcAddress"
        Exit Sub 'Exit here
    End If
    mh_ProcAddress = apiError 'Dump
    
    apiError = SetWindowsHookEx(WH_MOUSE, mh_ProcAddress, dllInstance, 0)
    If apiError = 0 Then
        Failed "SetWindowsHookEx"
        Exit Sub 'Exit here
    End If
    mh_Hook = apiError 'Dump
    
    MouseHookInit frmMain.picHM.hwnd, mh_Hook
End Sub

Public Sub UnInstallMouseHook()
    'Clean up hook
    If UnhookWindowsHookEx(mh_Hook) = 0 Then Failed "UnhookWindowsHookEx"
End Sub

Public Function Hooks_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        'Case WM_ACTIVATE 'Shell
        Dim POINTAPI As POINTAPI
        
        Case WM_MOUSEMOVE 'Mouse
            'GetCursorPos POINTAPI
            POINTAPI.X = LoWord(lParam)
            POINTAPI.Y = HiWord(lParam)
            
            If frmMain.mnuMouseMovOO.Checked = True Then
                'Calculates X
                If MouseMovTmpX > POINTAPI.X Then 'If placement of cursor has changed
                    MouseMovX = (MouseMovTmpX - POINTAPI.X) + MouseMovX 'Calculates difference between current pos and old pos
                    MouseMovTmpX = POINTAPI.X 'Resets the point of reference
                End If
                If MouseMovTmpX < POINTAPI.X Then 'If placement of cursor has changed
                    MouseMovX = (POINTAPI.X - MouseMovTmpX) + MouseMovX 'Calculates difference between current pos and old pos
                    MouseMovTmpX = POINTAPI.X 'Resets the point of reference
                End If
                
                'Calculates Y
                If MouseMovTmpY > POINTAPI.Y Then 'If placement of cursor has changed
                    MouseMovY = (MouseMovTmpY - POINTAPI.Y) + MouseMovY 'Calculates difference between current pos and old pos
                    MouseMovTmpY = POINTAPI.Y 'Resets the point of reference
                End If
                If MouseMovTmpY < POINTAPI.Y Then 'If placement of cursor has changed
                    MouseMovY = (POINTAPI.Y - MouseMovTmpY) + MouseMovY 'Calculates difference between current pos and old pos
                    MouseMovTmpY = POINTAPI.Y 'Resets the point of reference
                End If
                
                frmMouseMovement.txtX.Text = MouseMovX
                frmMouseMovement.txtY.Text = MouseMovY
                frmMouseMovement.txtTotal.Text = MouseMovX + MouseMovY
            End If
            
            If frmMain.mnuMouseWarpOO.Checked = True Then
                If POINTAPI.X = ScreenEdge.X - 1 Then  'If at right edge reset to left
                    SetCursorPos 1, POINTAPI.Y
                    MouseWarp = MouseWarp + 1 'Increments total
                Else
                    If POINTAPI.X = 0 Then 'If at left edge reset to right
                        SetCursorPos ScreenEdge.X - 2, POINTAPI.Y
                        MouseWarp = MouseWarp + 1 'Increments total
                    End If
                End If
                
                If POINTAPI.Y = ScreenEdge.Y - 1 Then 'If at bottom edge reset to top
                    SetCursorPos POINTAPI.X, 1
                    MouseWarp = MouseWarp + 1 'Increments total
                Else
                    If POINTAPI.Y = 0 Then 'If at top edge reset to bottom
                        SetCursorPos POINTAPI.X, ScreenEdge.Y - 2
                        MouseWarp = MouseWarp + 1 'Increments total
                    End If
                End If
                
                frmMouseWarp.txtTotal.Text = MouseWarp
            End If
            
        Case sh_RetMsg
            Select Case wParam
                'Case HSHELL_WINDOWCREATED
                'Case HSHELL_WINDOWDESTROYED
                'Case HSHELL_WINDOWACTIVATED
                'Case HSHELL_LANGUAGE
                'Case HSHELL_GETMINRECT
                Case HSHELL_REDRAW
                    If frmMain.mnuWindowKillerOO.Checked = True Then
                        'Errors are caused by changes in array while its going
                        On Error Resume Next
                        
                        If WindowKillerNum > 0 Then 'Cant send no messages
                            Dim tmpInt As Integer
                            Dim strWindowTitle As String
                            
                            strWindowTitle = Space$(512)
                            apiError = GetWindowText(lParam, strWindowTitle, 512)
                            strWindowTitle = Fix_NullTermStr(strWindowTitle)
                            
                            
                            For tmpInt = 1 To WindowKillerNum 'Cycle through array
                                If strWindowTitle = WindowKiller(tmpInt) Then
                                    'PostMessage lParam, WM_CLOSE, 0, 0
                                    apiError = PostMessage(lParam, WM_DESTROY, 0&, 0&)
                                    apiError = PostMessage(lParam, WM_NCDESTROY, 0&, 0&)
                                    'PostMessage tmpHandle, WM_MDIDESTROY, 0, 0

                                    Exit For
                                End If
                            Next tmpInt
                        Else
                            frmMain.mnuWindowKillerOO.Checked = False
                        End If
                    End If
                'Case HSHELL_TASKMAN
                'Case HSHELL_ACTIVATESHELLWINDOW
            End Select
        
        Case Else
            Hooks_Proc = DefWindowProc(frmMain.picHM.hwnd, uMsg, wParam, lParam)
    End Select
End Function

Public Sub InstallShellHook()
    apiError = RegisterWindowMessage(ByVal "SHELLHOOK")
    If apiError = 0 Then
        Failed "RegisterShellHook"
        Exit Sub
    End If
    sh_RetMsg = apiError
    
    Call RegisterShellHook(frmMain.picHM.hwnd, RSH_REGISTER)  ' Or RSH_REGISTER_TASKMAN Or RSH_REGISTER_PROGMAN)
End Sub

Public Sub UnInstallShellHook()
    If RegisterShellHook(frmMain.picHM.hwnd, RSH_DEREGISTER) = 0 Then Failed "RegisterShellHook"
End Sub

Public Sub StartHooks()
    'Load dll
    apiError = LoadLibrary(Dirs.System & "\kira_ext.dll")
    If apiError = 0 Then
        Failed "LoadLibraryEx"
        Exit Sub 'Exit here
    End If
    dllInstance = apiError 'Dump
    
    'Get old procedure
    apiError = GetWindowLong(frmMain.picHM.hwnd, GWL_WNDPROC)
    If apiError = 0 Then
        Failed "GetWindowLong"
        Exit Sub      'Exit here
    End If
    oldMH_Proc = apiError
    
    'Set new procedure
    If SetWindowLong(frmMain.picHM.hwnd, GWL_WNDPROC, AddressOf Hooks_Proc) = 0 Then Failed "SetWindowLong"
    
    
    InstallMouseHook
    InstallShellHook
End Sub

Public Sub StopHooks()
    If FreeLibrary(dllInstance) = 0 Then Failed "FreeLibrary"
    If SetWindowLong(frmMain.picHM.hwnd, GWL_WNDPROC, oldMH_Proc) = 0 Then Failed "SetWindowLong"
    
    UnInstallMouseHook
    UnInstallShellHook
End Sub
