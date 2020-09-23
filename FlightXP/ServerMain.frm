VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form ServerMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flight XP Server"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   2655
   End
   Begin RichTextLib.RichTextBox tMessages 
      Height          =   3855
      Left            =   3000
      TabIndex        =   12
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6800
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"ServerMain.frx":0000
   End
   Begin VB.CommandButton cStopServer 
      Caption         =   "Stop Server"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1800
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame fServerInfo 
      Caption         =   "Server Info"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
      Begin VB.Label lIP 
         Caption         =   "IP: ?.?.?.?"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lStart 
         Caption         =   "Start time: N/A"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lPlayers 
         Caption         =   "Players: 0/0"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.TextBox tCommand 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Frame fStartServer 
      Caption         =   "Start Server"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cStartServer 
         Caption         =   "Start Server"
         Height          =   615
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tMaxPlayers 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Text            =   "8"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox tConnectPort 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "8500"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Max. Players:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "ConnectPort:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin MSWinsockLib.Winsock SockU 
      Index           =   0
      Left            =   360
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock SockT 
      Index           =   0
      Left            =   720
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ConnectSock 
      Left            =   0
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "ServerMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cAbout_Click()
  Call MsgBox("Programmed by Jacob Poul Richardt a.k.a. Lillemanden" & vbCrLf & vbCrLf & "If you have a comment, drop me a line at: lilleman@fmk.kollegienet.dk", vbOKOnly + vbInformation, "About")
End Sub

Private Sub ConnectSock_ConnectionRequest(ByVal requestID As Long)
  ConnectSock.Close
  ConnectSock.Accept (requestID)
  If NAPlayers < NPlayers Then
    Dim Spot As Integer
    Spot = FreePlayers
    Players(Spot).Joining = True
    Players(Spot).InGame = True
    Players(Spot).Engine = False
    Players(Spot).Xploing = 0
    Players(Spot).Posi.X = 0
    Players(Spot).Posi.Y = 0
    Players(Spot).Direction = 0
    SockU(Spot).Close
    SockT(Spot).Close
    SockU(Spot).Bind
    SockT(Spot).Listen
    ConnectSock.SendData ConRoom & Spot
    ShowMes (ConnectSock.RemoteHostIP & " is connecting to playerslot " & Spot & ".")
    lPlayers.Caption = "Players: " & NFreeSlots & "/" & NPlayers
  Else
    ConnectSock.SendData ConFull
    ShowMes (ConnectSock.RemoteHostIP & " tried to connect (server full).")
  End If
  DoEvents
  ConnectSock.Close
  ConnectSock.Listen
End Sub

Private Sub cStartServer_Click()
  fStartServer.Enabled = False
  NPlayers = tMaxPlayers
  lPlayers = "Players: 0/" & NPlayers
  lStart = "Start time: " & Time
  fServerInfo.Enabled = True
  cStopServer.Enabled = True
  tCommand.Enabled = True
  tCommand.SetFocus
  ReadyPorts
  SLoop
End Sub

Private Sub cStopServer_Click()
  SendQuit (NPlayers)
  Running = False
  cStopServer.Enabled = False
  fServerInfo.Enabled = False
  tCommand.Enabled = False
  UnloadPorts
  fStartServer.Enabled = True
  ShowMes ("Server is no longer running.")
End Sub

Private Sub Form_Load()
  ServerMain.Caption = "Flight XP Server - ver. " & App.Major & "." & App.Minor & "." & App.Revision
  lIP = "IP: " & ConnectSock.LocalIP
  SetCodeColors
  Longside = 15 ^ 2 + 5 ^ 2
  If Dir(App.Path & "\SVSetting.FXP") = "SVSetting.FXP" Then
    Dim CSettings As Settings
    Open App.Path & "\SVSetting.FXP" For Random Access Read As #1 Len = Len(CSettings)
    Get #1, , CSettings
    Close #1
    ServerMain.tMaxPlayers = CSettings.Maxplayers
    ServerMain.tConnectPort = CSettings.Port
  Else
    ServerMain.tMaxPlayers = 8
    ServerMain.tConnectPort = 8500
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Running Then cStopServer_Click
  Dim CSettings As Settings
  CSettings.Maxplayers = ServerMain.tMaxPlayers
  CSettings.Port = ServerMain.tConnectPort
  Open App.Path & "\SVSetting.FXP" For Random Access Write As #1 Len = Len(CSettings)
  Put #1, , CSettings
  Close #1
  End
End Sub

Private Sub ReadyPorts()
  Dim I As Integer
  SockU(0).LocalPort = tConnectPort + 1
  SockT(0).LocalPort = tConnectPort + 2
  For I = 1 To NPlayers
    Load SockU(I)
    Load SockT(I)
    SockU(I).Close
    SockT(I).Close
    SockU(I).LocalPort = tConnectPort + 1 + I * 2
    SockT(I).LocalPort = tConnectPort + 2 + I * 2
  Next I
  ConnectSock.LocalPort = tConnectPort
  ConnectSock.Close
  ConnectSock.Listen
End Sub

Private Sub UnloadPorts()
  Dim I As Integer
  ConnectSock.Close
  SockU(0).Close
  SockT(0).Close
  For I = 1 To NPlayers
    SockU(I).Close
    SockT(I).Close
    Unload SockU(I)
    Unload SockT(I)
  Next I
End Sub

Private Sub SockU_ConnectionRequest(Index As Integer, ByVal requestID As Long)
  SockU(Index).Close
  SockU(Index).Accept requestID
  DoEvents
End Sub

Private Sub SockU_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  Dim Data As String
  SockU(Index).GetData Data
  Players(Index).Posi.X = (StrToInt(Mid(Data, 1, 2)) + 2 ^ 15) / 80
  Players(Index).Posi.Y = (StrToInt(Mid(Data, 3, 2)) + 2 ^ 15) / 80
  If StrToInt(Mid(Data, 5, 2)) < 0 Then
    Players(Index).Engine = True
    Players(Index).Direction = (StrToInt(Mid(Data, 5, 2)) + 2 ^ 15) / 5000
  Else
    Players(Index).Engine = False
    Players(Index).Direction = StrToInt(Mid(Data, 5, 2)) / 5000
  End If
  Exit Sub
End Sub

Private Sub SockT_ConnectionRequest(Index As Integer, ByVal requestID As Long)
  SockT(Index).Close
  SockT(Index).Accept requestID
  Dim JoinData As String
  JoinData = TJoinInfo & Chr(NPlayers) & Chr(NAPlayers)
  Dim I As Integer, P As Integer
  For I = 0 To NPlayers - 1
    If Players(I).InGame And Not (Players(I).Joining) Then
      JoinData = JoinData & Chr(I) & Players(I).Name & Chr(Players(I).Color And 255) & Chr((Players(I).Color And 65280) \ 256) & Chr((Players(I).Color And 16776960) \ 65536) & ActiveShots(I)
    End If
  Next I
  SockT(Index).SendData JoinData
  DoEvents
End Sub

Private Sub SockT_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  Dim Data As String
  SockT(Index).GetData Data
  Select Case Left(Data, 1)
    Case TJoinInfo
      Players(Index).Name = Mid(Data, 2, 25)
      Players(Index).Color = RGB(Asc(Mid(Data, 27, 1)), Asc(Mid(Data, 28, 1)), Asc(Mid(Data, 29, 1)))
      Call ShowMes(Trim(Players(Index).Name) & " has joined the game.", CPubMsg)
      PlayerJoins (Index)
      Players(Index).Joining = False
    Case TShot
      Dim CShot As Integer
      CShot = StrToInt(Mid(Data, 2, 2))
      Shots(Index, CShot).InGame = True
      NShots(Index) = NShots(Index) + 1
      Shots(Index, CShot).Posi.X = (StrToInt(Mid(Data, 4, 2)) + 2 ^ 15) / 80
      Shots(Index, CShot).Posi.Y = (StrToInt(Mid(Data, 6, 2)) + 2 ^ 15) / 80
      Shots(Index, CShot).Speed.X = StrToInt(Mid(Data, 8, 2)) / 30
      Shots(Index, CShot).Speed.Y = StrToInt(Mid(Data, 10, 2)) / 30
      Call SendShot(Index, CShot)
    Case TSay
      Call ShowMes(Trim(Players(Index).Name) & ": " & Mid(Data, 2), CPSpeak)
      Call SendSay(Index, Mid(Data, 2))
    Case TChangeInfo
      Call ShowMes(Trim(Players(Index).Name) & " changed playeroptions" & IIf(Players(Index).Name <> Mid(Data, 2, 25), ", and is now known as " & Trim(Mid(Data, 2, 25)) & ".", "."), CPubMsg)
      Players(Index).Name = Mid(Data, 2, 25)
      Players(Index).Color = RGB(Asc(Mid(Data, 27, 1)), Asc(Mid(Data, 28, 1)), Asc(Mid(Data, 29, 1)))
      SendChangeInfo (Index)
    Case TQuit
      Players(Index).InGame = False
      SockT(Index).Close
      SockU(Index).Close
      SendQuit (Index)
      Call ShowMes(Trim(Players(Index).Name) & " disconnected.", CPubMsg)
      lPlayers.Caption = "Players: " & NFreeSlots & "/" & NPlayers
    End Select
End Sub

Private Sub tCommand_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    tCommand = Replace(tCommand, vbCrLf, "")
    If tCommand <> "" Then
      If Left(tCommand, 1) = "\" Then
        ServerCommand (Mid(tCommand, 2))
      Else
        Call SendSay(NPlayers, tCommand)
        Call ShowMes("Server: " & tCommand, CSSpeak)
      End If
      tCommand = ""
    End If
  End If
End Sub

'Does Not Work

Public Sub SockU_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  SockU(Index).Close
  Players(Index).InGame = False
  SockT(Index).Close
  SendQuit (Index)
  Call ShowMes(Trim(Players(Index).Name) & "'s UDP socket caused an error (" & Description & "). Connection closed.")
  lPlayers.Caption = "Players: " & NFreeSlots & "/" & NPlayers
End Sub

Public Sub SockT_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  SockT(Index).Close
  Players(Index).InGame = False
  SockU(Index).Close
  SendQuit (Index)
  Call ShowMes(Trim(Players(Index).Name) & "'s TCP socket caused an error (" & Description & "). Connection closed.")
  lPlayers.Caption = "Players: " & NFreeSlots & "/" & NPlayers
End Sub

Public Sub ConnectSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  ConnectSock.Close
  ConnectSock.Listen
  Call ShowMes("The connection socket caused an error (" & Description & "). Socket has been reset.")
End Sub
