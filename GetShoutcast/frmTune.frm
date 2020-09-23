VERSION 5.00
Begin VB.Form frmTune 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tuner"
   ClientHeight    =   1410
   ClientLeft      =   4470
   ClientTop       =   3825
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3495
   Begin VB.CommandButton CmdCancel 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "8000"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "0.0.0.0"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblPort 
      Caption         =   "Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblServer 
      Caption         =   "Server IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmTune"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''
''    Module written by Fernando Aldea G.   ''
''    e-mail: fernando_aldea@terra.cl       ''
''    Release October, 2004                 ''
''                                          ''
''    sorry for not comment the code        ''
''    & sorry for my English!               ''
''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim sServer As String
    Dim sPath As String
    Dim pos As Integer
    
    'get out the http://
    sServer = Replace(Me.txtServer.Text, "http://", "")
    
    'split url
    pos = InStr(1, sServer, "/")
    If pos > 0 Then
        sPath = Mid(sServer, pos)
        sServer = Left(sServer, pos - 1)
    End If

    frmMain.Tune sServer, Me.txtPort.Text, sPath
    Unload Me
End Sub
