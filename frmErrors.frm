VERSION 5.00
Begin VB.Form frmErrors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Errors"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmErrors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtErrorNumber 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdGetInfo 
      Caption         =   "Get Info"
      Height          =   350
      Left            =   3120
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtDescription 
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label lblErrorNumber 
      Caption         =   "Error Number"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "frmErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetInfo_Click()
    txtDescription.Text = "" 'Clear
    If txtErrorNumber.Text = "" Then Exit Sub
    
    On Error Resume Next
    
    Dim errDescription As String
    
    apiError = CLng(txtErrorNumber.Text)
    txtErrorNumber.Text = apiError
    
    Errors.Errors apiError, "", errDescription, True
    
    txtDescription.Text = errDescription
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
