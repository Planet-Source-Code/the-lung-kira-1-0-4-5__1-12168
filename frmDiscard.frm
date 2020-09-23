VERSION 5.00
Begin VB.Form frmDiscard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discard"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "frmDiscard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3810
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
   Begin VB.CommandButton cmdSendData 
      Caption         =   "Send Data"
      Height          =   350
      Left            =   2640
      TabIndex        =   7
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtData 
      Height          =   2205
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1080
      Width           =   3495
   End
   Begin VB.ComboBox cboMethod 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   350
      Left            =   1680
      TabIndex        =   6
      Top             =   3360
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
   Begin VB.Label lblData 
      Caption         =   "Data"
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
Attribute VB_Name = "frmDiscard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngSocket As Long

Dim sockaddr As sockaddr

Private Sub cmdSendData_Click()
    cmdSendData.Enabled = False
    
    Select Case cboMethod.ListIndex
        Dim strBuffer As String
        
        Case 0 'UDP
            lngSocket = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP)
            If lngSocket = INVALID_SOCKET Then WinSockError "socket"

            With sockaddr
                .sin_addr = inet_addr(txtIP.Text & Chr(0))
                .sin_family = AF_INET
                .sin_port = htons(9)
                .sin_zero = String$(8, 0)
            End With
            
            strBuffer = txtData.Text
            
            If sendto(lngSocket, strBuffer, Len(strBuffer), 0, sockaddr, Len(sockaddr)) = SOCKET_ERROR Then WinSockError "sendto"
            
            If shutdown(lngSocket, 1) = SOCKET_ERROR Then WinSockError "shutdown"
            If closesocket(lngSocket) = SOCKET_ERROR Then WinSockError "closesocket"
        Case 1 'TCP
            lngSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
            If lngSocket = INVALID_SOCKET Then WinSockError "socket"
            
            With sockaddr
                .sin_addr = inet_addr(txtIP.Text & Chr(0))
                .sin_family = AF_INET
                .sin_port = htons(9)
                .sin_zero = String$(8, 0)
            End With
            
            strBuffer = txtData.Text
            
            If connect(lngSocket, sockaddr, Len(sockaddr)) = SOCKET_ERROR Then WinSockError "connect"
            If send(lngSocket, strBuffer, Len(strBuffer), 0) = SOCKET_ERROR Then WinSockError "send"
            
            If shutdown(lngSocket, 1) = SOCKET_ERROR Then WinSockError "shutdown"
            If closesocket(lngSocket) = SOCKET_ERROR Then WinSockError "closesocket"
    End Select
    
    cmdSendData.Enabled = True
End Sub

Private Sub cmdStop_Click()
    If shutdown(lngSocket, 1) = SOCKET_ERROR Then WinSockError "shutdown"
    If closesocket(lngSocket) = SOCKET_ERROR Then WinSockError "closesocket"
    
    cmdSendData.Enabled = True
End Sub

Private Sub Form_Load()
    With cboMethod
        .AddItem "UDP"
        .AddItem "TCP"
    End With
    
    cboMethod.ListIndex = GetSettingLong(HKEY_LOCAL_MACHINE, "Software\Kira\Discard", "Method")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettingLong HKEY_LOCAL_MACHINE, "Software\Kira\Discard", "Method", cboMethod.ListIndex
    
    'If they close make sure its cleaned up
    If shutdown(lngSocket, 1) = SOCKET_ERROR Then WinSockError "shutdown"
    If closesocket(lngSocket) = SOCKET_ERROR Then WinSockError "closesocket"
            
    Unload Me
End Sub
