VERSION 5.00
Begin VB.Form frmChecksum 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Checksum"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "frmChecksum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCRC32DEC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtAdler32DEC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtAdler32HEX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtCRC32HEX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose"
      Height          =   350
      Left            =   2640
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblAdler32HEX 
      Caption         =   "Adler32 HEX"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblAdler32DEC 
      Caption         =   "Adler32 DEC"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblCRC32DEC 
      Caption         =   "CRC32 DEC"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblCRC32HEX 
      Caption         =   "CRC32 HEX"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmChecksum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFileName As String

Private Sub cmdChoose_Click()
    GetOpenName hwnd, "Open", strFileName
    
    'Error checking
    If strFileName = "" Then
        Exit Sub 'Dont worry just exit
    End If
    
    Dim strFileContents As String
    Open strFileName For Binary As #1 'Opens it for binary
        strFileContents = Space$(LOF(1)) 'Pads to length of string
        Get #1, , strFileContents 'Dumps contents of file to string
    Close #1
    
    Dim crc As Long
    Dim adler As Long
    
    crc = crc32(crc, strFileContents, Len(strFileContents))
    adler = adler32(adler, strFileContents, Len(strFileContents))
    
    txtCRC32HEX.Text = Hex$(crc)
    txtCRC32DEC.Text = crc
    txtAdler32HEX.Text = Hex$(adler)
    txtAdler32DEC.Text = adler
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
