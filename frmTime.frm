VERSION 5.00
Begin VB.Form frmTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "frmTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtReturned 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.ComboBox cboMethod 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.PictureBox picTime 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdGetData 
      Caption         =   "Get Data"
      Height          =   350
      Left            =   2640
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   350
      Left            =   1680
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblIP 
      Caption         =   "IP"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblReturned 
      Caption         =   "Returned"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblMethod 
      Caption         =   "Method"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngSocket As Long

Private Sub cmdGetData_Click()
    cmdGetData.Enabled = False
    txtReturned.Text = ""
    
    Dim sockaddr As sockaddr
    
    Select Case cboMethod.ListIndex
        Dim strBuffer As String * 8
        
        Case 0 'UDP
            lngSocket = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP)
            If lngSocket = INVALID_SOCKET Then WinSockError "socket"
            
            If WSAAsyncSelect(lngSocket, picTime.hwnd, ByVal WM_MOUSEMOVE, FD_CLOSE Or FD_READ) = SOCKET_ERROR Then WinSockError "WSAAsyncSelect"
            
            With sockaddr
                .sin_addr = inet_addr(txtIP.Text & Chr(0))
                .sin_family = AF_INET
                .sin_port = htons(37)
                .sin_zero = String$(8, 0)
            End With
            
            If sendto(lngSocket, 0&, 0, 0, sockaddr, Len(sockaddr)) = SOCKET_ERROR Then WinSockError "sendto"
        Case 1 'TCP
            lngSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
            If lngSocket = INVALID_SOCKET Then WinSockError "socket"
            
            With sockaddr
                .sin_addr = inet_addr(txtIP.Text & Chr(0))
                .sin_family = AF_INET
                .sin_port = htons(37)
                .sin_zero = String$(8, 0)
            End With
            
            If connect(lngSocket, sockaddr, Len(sockaddr)) = SOCKET_ERROR Then WinSockError "connect"
            If WSAAsyncSelect(lngSocket, picTime.hwnd, ByVal WM_MOUSEMOVE, FD_CLOSE Or FD_READ) = SOCKET_ERROR Then WinSockError "WSAAsyncSelect"
    End Select
End Sub

Private Sub cmdStop_Click()
    If shutdown(lngSocket, 1) = SOCKET_ERROR Then WinSockError "shutdown"
    If closesocket(lngSocket) = SOCKET_ERROR Then WinSockError "closesocket"
    
    cmdGetData.Enabled = True
End Sub

Private Sub Form_Load()
    With cboMethod
        .AddItem "UDP"
        .AddItem "TCP"
    End With
    
    cboMethod.ListIndex = GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira\Time", "Method")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\Time", "Method", cboMethod.ListIndex
    
    'If they close make sure its cleaned up
    If shutdown(lngSocket, 1) = SOCKET_ERROR Then WinSockError "shutdown"
    If closesocket(lngSocket) = SOCKET_ERROR Then WinSockError "closesocket"
    
    Unload Me
End Sub

Private Sub picTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case cboMethod.ListIndex
        Dim strBuffer As String * 8
        Dim dblTime As Double
        Dim strData As String
        Dim sockaddr As sockaddr
        
        Case 0 'UDP
            apiError = recvfrom(lngSocket, ByVal strBuffer, Len(strBuffer), 0, sockaddr, Len(sockaddr))
            If apiError = SOCKET_ERROR Then WinSockError "recvfrom"
            
            If apiError > 0 Then
                strData = Left$(strBuffer, apiError)
                
                If Len(strData) = 4 Then
                    dblTime = Str2Lng(strData)
                    If dblTime < 0 Then dblTime = (2 ^ 32) + dblTime
                    
                    txtReturned.Text = DateAdd("s", dblTime - 2208988800#, "1/1/1970")
                End If
            End If
            
            If shutdown(lngSocket, 1) = SOCKET_ERROR Then WinSockError "shutdown"
            If closesocket(lngSocket) = SOCKET_ERROR Then WinSockError "closesocket"
        Case 1 'TCP
            apiError = recv(lngSocket, ByVal strBuffer, Len(strBuffer), 0)
            If apiError = SOCKET_ERROR Then WinSockError "recv"
            
            If apiError > 0 Then
                strData = Left$(strBuffer, apiError)
                
                If Len(strData) = 4 Then
                    dblTime = Str2Lng(strData)
                    If dblTime < 0 Then dblTime = (2 ^ 32) + dblTime
                    
                    txtReturned.Text = DateAdd("s", dblTime - 2208988800#, "1/1/1970")
                End If
            End If
            
            If shutdown(lngSocket, 1) = SOCKET_ERROR Then WinSockError "shutdown"
            If closesocket(lngSocket) = SOCKET_ERROR Then WinSockError "closesocket"
    End Select
    
    cmdGetData.Enabled = True
End Sub
