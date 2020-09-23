VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GetShoutcast by Fernando Aldea"
   ClientHeight    =   3600
   ClientLeft      =   3075
   ClientTop       =   2700
   ClientWidth     =   6150
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6150
   Begin VB.CheckBox CmdRecord 
      Caption         =   "RECORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   720
      Width           =   1095
   End
   Begin VB.Frame fmeOptions 
      Caption         =   "Record options:"
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   5895
      Begin VB.CheckBox chkFilebysong 
         Caption         =   "Create mp3 file by song (file name = title)"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   5415
      End
      Begin VB.TextBox txtOutPath 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Text            =   "c:\"
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Output Path:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdTune 
      Caption         =   "TUNE"
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   8
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   7
         Top             =   1080
         Width           =   285
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3120
         TabIndex        =   6
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "REC"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   0
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblBitrate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bitrate:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   480
         Width           =   345
      End
      Begin VB.Label lblRadio 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Radio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4575
      End
   End
   Begin MSWinsockLib.Winsock sckReceiver 
      Left            =   6000
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckSender 
      Left            =   6600
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblWinampStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Winamp Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   5895
   End
End
Attribute VB_Name = "frmMain"
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


Const ReqHeader = "" & _
                "GET $ HTTP/1.0" & vbCrLf & _
                "Host: %" & vbCrLf & _
                "User-Agent: WinampMPEG/2.7" & vbCrLf & _
                "Accept: */*" & vbCrLf & _
                "Icy-MetaData:1" & vbCrLf & _
                "Connection: close" & vbCrLf & vbCrLf
                

Dim WinampRequest As Boolean
Dim WinampConnected As Boolean
Dim IcyReceived As Boolean
Dim IcySended As Boolean
Dim sIcyHeader As String
Dim sData As String
Dim sMeta As String
Dim DataLen As Long
Dim MetaLen As Long
Dim nData As Long
Dim bMeta As Boolean
Dim iFile As Integer
Dim FileLen As Long
Dim StartTime As Date
Dim sPath As String  'UPDATE Feb 17


Public Sub StartServer()
    Me.sckSender.Close
    Me.sckSender.Protocol = sckTCPProtocol
    Me.sckSender.LocalPort = 8000
    Me.sckSender.Listen
    
    Me.lblWinampStatus.Caption = "Waiting for Winamp player open url '127.0.0.1:8000'..."

End Sub

Public Function DownServer()
    Me.sckSender.Close
    WinampRequest = False
    WinampConnected = False
End Function


Public Sub Tune(ByVal ServerIP As String, ByVal Port As Long, Optional Path As String)
    DownServer
    
    IcyReceived = False
    IcySended = False
    sIcyHeader = ""
    sData = ""
    sMeta = ""
    DataLen = 0
    MetaLen = 0
    nData = 0
    bMeta = False
    iFile = 0
    FileLen = 0
    sPath = IIf(Path = "", "/", Path)
    
    StartServer
    Me.sckReceiver.Close
    Me.sckReceiver.Connect ServerIP, Port
End Sub


Private Sub CmdRecord_Click()
    If Me.CmdRecord.Value = vbChecked Then
        If CreateOutFile Then
            Me.CmdRecord.Caption = "STOP"
            Me.lblStatus.Visible = True
            Me.lblFile.Enabled = True
            Me.lblSize.Enabled = True
            Me.lblTime.Enabled = True
        Else
            Me.CmdRecord.Value = vbUnchecked
        End If
    Else
        CloseOutFile
        Me.CmdRecord.Caption = "RECORD"
        Me.lblStatus.Visible = False
        Me.lblFile.Enabled = False
        Me.lblSize.Enabled = False
        Me.lblTime.Enabled = False
    End If
End Sub

Private Sub cmdTune_Click()
    frmTune.Show
End Sub

Private Sub Form_Load()
    Me.lblStatus.Visible = False
    Me.lblFile.Enabled = False
    Me.lblSize.Enabled = False
    Me.lblTime.Enabled = False
End Sub

Private Sub sckReceiver_Close()
    Me.lblRadio.Caption = "Disconnected"
End Sub
Private Sub sckReceiver_Connect()
    Dim sRequest As String
    
    Me.lblRadio.Caption = "Connected"
    sRequest = Replace$(ReqHeader, "%", Me.sckReceiver.RemoteHostIP)
    sRequest = Replace$(sRequest, "$", sPath)
    Me.sckReceiver.SendData sRequest
End Sub
Private Sub sckReceiver_ConnectionRequest(ByVal requestID As Long)
    Me.lblRadio.Caption = "Connecting ..."
End Sub
Private Sub sckReceiver_DataArrival(ByVal bytesTotal As Long)
    Dim pos As Long, pos2 As Long
    Dim sBuffer As String, seconds As String
    
    Me.sckReceiver.GetData sBuffer, , bytesTotal
    
    If Not IcyReceived Then
        sData = sData & sBuffer
        pos = InStr(1, sData, vbCrLf & vbCrLf)
        
        If pos > 0 Then
            If InStr(1, sData, "ICY 200 OK") Then
                sIcyHeader = Left(sData, pos + Len(vbCrLf & vbCrLf) - 1)
                'seek metaint
                pos = InStr(1, sData, "icy-metaint:") + Len("icy-metaint:")
                pos2 = InStr(pos, sData, vbCrLf)
                DataLen = Mid(sData, pos, pos2 - pos + 1)
                
                sBuffer = Mid(sData, Len(sIcyHeader) + 1)
                ShowInfo sIcyHeader
                IcyReceived = True
                bMeta = False
                GoSub SendFirstWinamp
            End If
        Else
            'some time out waiting for Icy header??
        End If
        
    End If
    
    
    If IcyReceived Then
        If WinampConnected Then
            Me.sckSender.SendData sBuffer
        End If
        
        While sBuffer <> ""
            sBuffer = ProcessBuffer(sBuffer, bMeta)
            GoSub SendFirstWinamp
        Wend
        
        If iFile Then
            seconds = DateDiff("s", StartTime, Now)
            Me.lblSize.Caption = "Size: " & (FileLen \ 1024) & " kb."
            Me.lblTime.Caption = "Time: " & (seconds \ 60) & ":" & Format((seconds Mod 60), "0#")
        End If
    End If
    
    Exit Sub
    
SendFirstWinamp:
            'send data to winamp first time?
            If bMeta = False And WinampRequest Then
                Me.sckSender.SendData sBuffer
                WinampRequest = False
                WinampConnected = True
            End If
            Return
    
End Sub


Function ProcessBuffer(ByVal sBuffer As String, ByRef esMeta As Boolean) As String
    Dim Remain As Long
    
        'incoming Buffer is data
        If esMeta = False Then
            Remain = DataLen - nData
            If (Remain <= Len(sBuffer)) Then
                nData = nData + Remain
                Call WriteOutFile(Left(sBuffer, Remain))
                nData = 0
                esMeta = True
                ProcessBuffer = Mid(sBuffer, Remain + 1)
            Else
                nData = nData + Len(sBuffer)
                Call WriteOutFile(sBuffer)
                ProcessBuffer = ""
            End If
                    
        'incoming buffer is metadata
        Else
            If MetaLen = 0 Then
                'get length of metadata (first byte of block * 16)
                MetaLen = Asc(Left(sBuffer, 1)) * 16
            End If
            
            Remain = MetaLen - Len(sMeta)
            If Remain = 0 Then
                esMeta = False
                ProcessBuffer = Mid(sBuffer, 2)
            ElseIf Remain <= Len(sBuffer) Then
                sMeta = sMeta & Mid(sBuffer, 2, Remain)
                
                ShowTitle sMeta
                
                sMeta = ""
                MetaLen = 0
                esMeta = False
                ProcessBuffer = Mid(sBuffer, Remain + 2)
            Else
                sMeta = sMeta & sBuffer
                ProcessBuffer = ""
            End If
        End If
        
End Function



Private Sub sckReceiver_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Me.lblRadio.Caption = "Error ocurred"
End Sub

Private Sub sckSender_Close()
    WinampConnected = False
    Me.lblWinampStatus.Caption = "Winamp Disconnected"
    StartServer
End Sub

Private Sub sckSender_ConnectionRequest(ByVal requestID As Long)
    Me.lblWinampStatus.Caption = "Winamp Connected"
    Me.sckSender.Close
    Me.sckSender.Accept requestID
End Sub

Private Sub sckSender_DataArrival(ByVal bytesTotal As Long)
    
    If IcyReceived Then
        Me.sckSender.SendData sIcyHeader
        WinampRequest = True
    End If
End Sub


Sub ShowTitle(ByVal sMetaData As String)
    Dim pos As Long, pos2 As Long
    
    'sTitle = Split(Split(sMetaData, ";")(0), "=")(1)
    
    'seek TitleSOng
    pos = InStr(1, sMetaData, "StreamTitle=") + Len("StreamTitle=")
    If pos > 0 Then
        pos2 = InStr(pos, sMetaData, ";")
        Me.lblTitle.Caption = "Title: " & Mid(sMetaData, pos, pos2 - pos)
    Else
        Me.lblTitle.Caption = "Title: " & "Title not Available"
    End If
    
    
    'Me.lblTitle.Caption = "Title: " & sTitle
    
    If Me.CmdRecord.Value = vbChecked And chkFilebysong.Value = vbChecked Then
        CloseOutFile
        CreateOutFile
    End If
    
End Sub

Sub ShowInfo(ByVal sIcy As String)
    Dim pos As Long, pos2 As Long
    
    'seek station name
    pos = InStr(1, sIcy, "icy-name:") + Len("icy-name:")
    If pos > 0 Then
        pos2 = InStr(pos, sIcy, vbCrLf)
        Me.lblRadio.Caption = Mid(sIcy, pos, pos2 - pos + 1)
    Else
        Me.lblRadio.Caption = "unknown station"
    End If
    
    'seek bit rate
    pos = InStr(1, sIcy, "icy-br:") + Len("icy-br:")
    If pos > 0 Then
        pos2 = InStr(pos, sIcy, vbCrLf)
        Me.lblBitrate.Caption = "Bitrate: " & Mid(sIcy, pos, pos2 - pos) & " Kbps"
    Else
        Me.lblBitrate.Caption = "Bitrate: " & "unknown bitrate"
    End If
    
End Sub

Function CreateOutFile() As Boolean
    Dim sPath As String
    Dim sFile As String
    
    If iFile <> 0 Then CloseOutFile
    
    FileLen = 0
    StartTime = Now
    If chkFilebysong.Value = vbUnchecked Then
        sFile = "File" & CLng(Timer) & ".mp3"
        sFile = InputBox("Enter file name: ", "New mp3 file", sFile)
        If sFile = "" Then Exit Function
    Else
        sFile = LTrim(Split(Me.lblTitle.Caption, ":")(1)) & ".mp3"
    End If
    
    iFile = FreeFile()
    sPath = Me.txtOutPath.Text
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    Open sPath & sFile For Binary As #iFile
    
    Me.lblFile.Caption = "File: " & sFile
    
    CreateOutFile = True

End Function

Sub WriteOutFile(ByVal sBuff As String)
    If iFile = 0 Then Exit Sub
    Put #iFile, , sBuff
    
    FileLen = FileLen + Len(sBuff)
    

End Sub

Sub CloseOutFile()
    Close #iFile
    iFile = 0
End Sub

