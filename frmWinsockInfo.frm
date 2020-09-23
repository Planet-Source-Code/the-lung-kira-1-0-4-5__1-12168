VERSION 5.00
Begin VB.Form frmWinsockInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Winsock Info"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmWinsockInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6000
      Width           =   6615
   End
   Begin VB.ListBox lstHosts 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   6615
   End
   Begin VB.ListBox lstNetworks 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   6615
   End
   Begin VB.ListBox lstServices 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   6615
   End
   Begin VB.ListBox lstProtocols 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   6615
   End
   Begin VB.TextBox txtWinsockSystemStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox txtWinsockDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblHosts 
      Caption         =   "Hosts"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblNetworks 
      Caption         =   "Networks"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblServices 
      Caption         =   "Services"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label lblProtocols 
      Caption         =   "Protocols"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblWinsockDescription 
      Caption         =   "Winsock Description"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblWinsockSystemStatus 
      Caption         =   "Winsock System Status"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmWinsockInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    txtWinsockDescription.Text = WinsockData.Description
    txtWinsockSystemStatus.Text = WinsockData.SystemStatus
    
    Dim strFileContents As String
    
    If WinID = "WIN32_WINDOWS" Then '9x
        'HOSTS
        If Dir$(Dirs.Windows & "\HOSTS") <> "" Then
            Open Dirs.Windows & "\HOSTS" For Input As #1
                Do While Not EOF(1) 'Loop until end of file
                    Line Input #1, strFileContents 'Read line into variable
                    
                    If Left$(strFileContents, 1) <> "#" Then
                    If Left$(strFileContents, 1) <> "" Then
                        lstServices.AddItem strFileContents
                    End If
                    End If
                Loop
            Close #1
        End If
        
        'NETWORKS
        If Dir$(Dirs.Windows & "\NETWORKS") <> "" Then
            strFileContents = ""
            
            Open Dirs.Windows & "\NETWORKS" For Input As #1
                Do While Not EOF(1) 'Loop until end of file
                    Line Input #1, strFileContents 'Read line into variable
                    
                    If Left$(strFileContents, 1) <> "#" Then
                    If Left$(strFileContents, 1) <> "" Then
                        lstNetworks.AddItem strFileContents
                    End If
                    End If
                Loop
            Close #1
        End If
        
        'PROTOCOL
        If Dir$(Dirs.Windows & "\PROTOCOL") <> "" Then
            strFileContents = ""
            
            Open Dirs.Windows & "\PROTOCOL" For Input As #1
                Do While Not EOF(1) 'Loop until end of file
                    Line Input #1, strFileContents 'Read line into variable
                    
                    If Left$(strFileContents, 1) <> "#" Then
                    If Left$(strFileContents, 1) <> "" Then
                        lstProtocols.AddItem strFileContents
                    End If
                    End If
                Loop
            Close #1
        End If
        
        'SERVICES
        If Dir$(Dirs.Windows & "\SERVICES") <> "" Then
            strFileContents = ""
            
            Open Dirs.Windows & "\SERVICES" For Input As #1
                Do While Not EOF(1) 'Loop until end of file
                    Line Input #1, strFileContents 'Read line into variable
                    
                    If Left$(strFileContents, 1) <> "#" Then
                    If Left$(strFileContents, 1) <> "" Then
                        lstServices.AddItem strFileContents
                    End If
                    End If
                Loop
            Close #1
        End If
    Else 'NT
        'HOSTS
        If Dir$(Dirs.System & "\drivers\etc\HOSTS") <> "" Then
            Open Dirs.System & "\drivers\etc\HOSTS" For Input As #1
                Do While Not EOF(1) 'Loop until end of file
                    Line Input #1, strFileContents 'Read line into variable
                    
                    If Left$(strFileContents, 1) <> "#" Then
                    If Left$(strFileContents, 1) <> "" Then
                        lstHosts.AddItem strFileContents
                    End If
                    End If
                Loop
            Close #1
        End If
        
        'NETWORKS
        If Dir$(Dirs.System & "\drivers\etc\NETWORKS") <> "" Then
            strFileContents = ""
            
            Open Dirs.System & "\drivers\etc\NETWORKS" For Input As #1
                Do While Not EOF(1) 'Loop until end of file
                    Line Input #1, strFileContents 'Read line into variable
                    
                    If Left$(strFileContents, 1) <> "#" Then
                    If Left$(strFileContents, 1) <> "" Then
                        lstNetworks.AddItem strFileContents
                    End If
                    End If
                Loop
            Close #1
        End If
        
        'PROTOCOL
        If Dir$(Dirs.System & "\drivers\etc\PROTOCOL") <> "" Then
            strFileContents = ""
            
            Open Dirs.System & "\drivers\etc\PROTOCOL" For Input As #1
                Do While Not EOF(1) 'Loop until end of file
                    Line Input #1, strFileContents 'Read line into variable
                    
                    If Left$(strFileContents, 1) <> "#" Then
                    If Left$(strFileContents, 1) <> "" Then
                        lstProtocols.AddItem strFileContents
                    End If
                    End If
                Loop
            Close #1
        End If
        
        'SERVICES
        If Dir$(Dirs.System & "\drivers\etc\SERVICES") <> "" Then
            strFileContents = ""
            
            Open Dirs.System & "\drivers\etc\SERVICES" For Input As #1
                Do While Not EOF(1) 'Loop until end of file
                    Line Input #1, strFileContents 'Read line into variable
                    
                    If Left$(strFileContents, 1) <> "#" Then
                    If Left$(strFileContents, 1) <> "" Then
                        lstServices.AddItem strFileContents
                    End If
                    End If
                Loop
            Close #1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstHosts_Click()
    txtData.Text = lstHosts.List(lstHosts.ListIndex)
End Sub

Private Sub lstNetworks_Click()
    txtData.Text = lstNetworks.List(lstNetworks.ListIndex)
End Sub

Private Sub lstProtocols_Click()
    txtData.Text = lstProtocols.List(lstProtocols.ListIndex)
End Sub

Private Sub lstServices_Click()
    txtData.Text = lstServices.List(lstServices.ListIndex)
End Sub
