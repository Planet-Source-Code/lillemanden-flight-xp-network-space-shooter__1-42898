VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flight XP"
   ClientHeight    =   9000
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Game 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9000
      Left            =   0
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin MSWinsockLib.Winsock ConnectSock 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock SockT 
         Left            =   960
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock SockU 
         Left            =   480
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
   End
   Begin VB.Menu mJoin 
      Caption         =   "Join"
   End
   Begin VB.Menu mPlayerOption 
      Caption         =   "Player Options"
   End
   Begin VB.Menu mGameOptions 
      Caption         =   "Game Options"
   End
   Begin VB.Menu mMessage 
      Caption         =   "Toggle Messagewindow"
      Enabled         =   0   'False
   End
   Begin VB.Menu mScore 
      Caption         =   "Toggle Scorewindow"
      Enabled         =   0   'False
   End
   Begin VB.Menu mAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ConnectSock_DataArrival(ByVal bytesTotal As Long)
  Dim Answer As String
  ConnectSock.GetData Answer
  Select Case Left(Answer, 1)
    Case ConFull
      Call ShowMes("The server is full.")
    Case ConRoom
      Call ShowMes("Connection accepted.")
      LocPlayerN = Mid(Answer, 2)
      SockU.Close
      SockT.Close
      SockU.RemoteHost = JoinForm.JoinIP
      SockT.RemoteHost = JoinForm.JoinIP
      SockU.RemotePort = JoinForm.JoinPort + 1 + 2 * LocPlayerN
      SockT.RemotePort = JoinForm.JoinPort + 2 + 2 * LocPlayerN
      SockU.Connect
      SockT.Connect
      DoEvents
  End Select
  ConnectSock.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Playing Then Disconnect
  Dim CSettings As Settings
  CSettings.Color = LocPColor
  CSettings.Name = LocPName
  CSettings.JoinHost = JoinForm.JoinIP.Text
  CSettings.JoinPort = JoinForm.JoinPort.Text
  CSettings.EngineEffect = NThrustFire
  Open App.Path & "\CLSetting.FXP" For Random Access Write As #1 Len = Len(CSettings)
  Put #1, , CSettings
  Close #1
  End
End Sub

Private Sub Form_Load()
  Main.Caption = "Flight XP - ver. " & App.Major & "." & App.Minor & "." & App.Revision
  SetCodeColors
  LocPColor = 75
  LocPName = "Player"
  Main.Top = 0
  Main.Left = 0
  fMessage.Top = 0
  fMessage.Left = Main.Width
  fMessage.Show
  If Dir(App.Path & "\CLSetting.FXP") = "CLSetting.FXP" Then
    Dim CSettings As Settings
    Open App.Path & "\CLSetting.FXP" For Random Access Read As #1 Len = Len(CSettings)
    Get #1, , CSettings
    Close #1
    LocPColor = CSettings.Color
    LocPName = CSettings.Name
    JoinForm.JoinIP.Text = CSettings.JoinHost
    JoinForm.JoinPort.Text = CSettings.JoinPort
    NThrustFire = CSettings.EngineEffect
  Else
    LocPColor = 75
    LocPName = "Player"
    JoinForm.JoinIP.Text = Main.ConnectSock.LocalIP
    JoinForm.JoinPort.Text = "8500"
    NThrustFire = 300
  End If
  DoEvents
  Main.Show
  Main.SetFocus
End Sub

Private Sub Game_LostFocus()
  HasFocus = False
End Sub

Private Sub Game_GotFocus()
  HasFocus = True
End Sub

Private Sub mAbout_Click()
  Call MsgBox("Programmed by Jacob Poul Richardt a.k.a. Lillemanden" & vbCrLf & vbCrLf & "If you have a comment, drop me a line at: lilleman@fmk.kollegienet.dk", vbOKOnly + vbInformation, "About")
End Sub

Private Sub mGameOptions_Click()
  Gameoptions.Show
End Sub

Private Sub mJoin_Click()
  If Playing Then
    Disconnect
  Else
    JoinForm.Show
  End If
End Sub

Private Sub mMessage_Click()
  If fMessage.Visible Then
    fMessage.Hide
  Else
    fMessage.Show
  End If
End Sub

Private Sub mPlayerOption_Click()
  PlayerOptions.Show
End Sub

Private Sub Disconnect(Optional SendQuit As Boolean = True)
  If Playing Then
    Call ShowMes("Disconnected.")
    Playing = False
    mJoin.Caption = "Join"
    mMessage.Enabled = False
    mScore.Enabled = False
    fMessage.SayText.Enabled = False
    fMessage.SayText.Text = ""
    If SendQuit And SockT.State = sckConnected Then SockT.SendData TQuit
    DoEvents
    SockT.Close
    SockU.Close
    ConnectSock.Close
  End If
End Sub

Private Sub SockT_DataArrival(ByVal bytesTotal As Long)
  Dim Data As String, CPlayer As Integer, CShot As Integer
  SockT.GetData Data
  Select Case Left(Data, 1)
    Case TJoinInfo
      NPlayers = Asc(Mid(Data, 2, 1))
      ReDim Players(NPlayers - 1)
      ReDim Shots(NPlayers - 1, MaxShots - 1)
      ReDim NShots(NPlayers - 1)
      ReDim Xplos(NPlayers - 1)
      If NThrustFire > 0 Then ReDim ThrustFire(NPlayers - 1, NThrustFire - 1)
      Players(LocPlayerN).Color = LocPColor
      Players(LocPlayerN).Name = LocPName
      Players(LocPlayerN).Ingame = True
      Dim P As Integer, RSpot As Long, NPShot As Integer
      RSpot = 4
      For P = 1 To Asc(Mid(Data, 3, 1))
        CPlayer = Asc(Mid(Data, RSpot, 1))
        Players(CPlayer).Ingame = True
        RSpot = RSpot + 1
        Players(CPlayer).Name = Mid(Data, RSpot, 25)
        RSpot = RSpot + 25
        Players(CPlayer).Color = RGB(Asc(Mid(Data, RSpot, 1)), Asc(Mid(Data, RSpot + 1, 1)), Asc(Mid(Data, RSpot + 2, 1)))
        RSpot = RSpot + 3
        NShots(CPlayer) = StrToInt(Mid(Data, RSpot, 2))
        RSpot = RSpot + 2
        For NPShot = 1 To NShots(CPlayer)
          CShot = StrToInt(Mid(Data, RSpot, 2))
          Shots(CPlayer, CShot).Ingame = True
          Shots(CPlayer, CShot).Posi.X = (StrToInt(Mid(Data, RSpot + 2, 2)) + 2 ^ 15) / 80
          Shots(CPlayer, CShot).Posi.Y = (StrToInt(Mid(Data, RSpot + 4, 2)) + 2 ^ 15) / 80
          Shots(CPlayer, CShot).Speed.X = StrToInt(Mid(Data, RSpot + 6, 2)) / 30
          Shots(CPlayer, CShot).Speed.Y = StrToInt(Mid(Data, RSpot + 8, 2)) / 30
          RSpot = RSpot + 10
        Next NPShot
      Next P
      DoEvents
      SockT.SendData TJoinInfo & LocInfo
      DoEvents
      GLoop
    Case TJoin
      CPlayer = Asc(Mid(Data, 2, 1))
      Players(CPlayer).Name = Mid(Data, 3, 25)
      Players(CPlayer).Color = RGB(Asc(Mid(Data, 28, 1)), Asc(Mid(Data, 29, 1)), Asc(Mid(Data, 30, 1)))
      Players(CPlayer).Ingame = True
      Call ShowMes(Trim(Players(CPlayer).Name) & " joined the game.", CPubMsg)
    Case TShot
      CPlayer = Asc(Mid(Data, 2, 1))
      CShot = StrToInt(Mid(Data, 3, 2))
      Shots(CPlayer, CShot).Ingame = True
      NShots(CPlayer) = NShots(CPlayer) + 1
      Shots(CPlayer, CShot).Posi.X = (StrToInt(Mid(Data, 5, 2)) + 2 ^ 15) / 80
      Shots(CPlayer, CShot).Posi.Y = (StrToInt(Mid(Data, 7, 2)) + 2 ^ 15) / 80
      Shots(CPlayer, CShot).Speed.X = StrToInt(Mid(Data, 9, 2)) / 30
      Shots(CPlayer, CShot).Speed.Y = StrToInt(Mid(Data, 11, 2)) / 30
    Case TSay
      CPlayer = Asc(Mid(Data, 2, 1))
      If CPlayer = NPlayers Then
        Call ShowMes("Server: " & Mid(Data, 3), CSSpeak)
      Else
        Call ShowMes(Trim(Players(CPlayer).Name) & ": " & Mid(Data, 3), CPSpeak)
      End If
    Case TChangeInfo
      CPlayer = Asc(Mid(Data, 2, 1))
      Call ShowMes(Trim(Players(CPlayer).Name) & " changed playeroptions" & IIf(Players(CPlayer).Name <> Mid(Data, 3, 25), ", and is now known as " & Trim(Mid(Data, 3, 25)) & ".", "."), CPubMsg)
      Players(CPlayer).Name = Mid(Data, 3, 25)
      Players(CPlayer).Color = RGB(Asc(Mid(Data, 28, 1)), Asc(Mid(Data, 29, 1)), Asc(Mid(Data, 30, 1)))
    Case TQuit
      CPlayer = Asc(Mid(Data, 2, 1))
      If CPlayer = NPlayers Then
        Call ShowMes("The server quit.", CPubMsg)
        Disconnect
      Else
        Players(CPlayer).Ingame = False
        Call ShowMes(Trim(Players(CPlayer).Name) & " disconnected.", CPubMsg)
      End If
    Case TKick
      CPlayer = Asc(Mid(Data, 2, 1))
      If CPlayer = LocPlayerN Then
        Call ShowMes("You have been kicked!", CPubMsg)
        Disconnect (False)
      Else
        Call ShowMes(Trim(Players(CPlayer).Name) & " has been kicked.", CPubMsg)
        Players(CPlayer).Ingame = False
      End If
    Case TKill
      Dim KPlayer As Integer
      Dim KShot As Integer
      CPlayer = Asc(Mid(Data, 3, 1))
      KPlayer = Asc(Mid(Data, 4, 1))
      KShot = StrToInt(Mid(Data, 5, 2))
      Players(CPlayer).Xploing = DeadTime
      SetXplo (CPlayer)
      Select Case Mid(Data, 2, 1)
        Case KillShot
          Shots(KPlayer, KShot).Ingame = False
          NShots(KPlayer) = NShots(KPlayer) - 1
          If CPlayer = KPlayer Then
            If CPlayer = LocPlayerN Then
              LocPSpeed.X = 0
              LocPSpeed.Y = 0
              Randomize
              Players(LocPlayerN).Direction = Rnd() * 6.283185307
              Players(LocPlayerN).Posi.X = Rnd() * 800
              Players(LocPlayerN).Posi.Y = Rnd() * 600
              Call ShowMes("You killed yourself!", CAction)
            Else
              Call ShowMes(Trim(Players(CPlayer).Name) & " committed suicide!", CAction)
            End If
          Else
            If KPlayer = LocPlayerN Then
              Call ShowMes("You killed " & Trim(Players(CPlayer).Name) & ".", CAction)
            ElseIf CPlayer = LocPlayerN Then
              LocPSpeed.X = 0
              LocPSpeed.Y = 0
              Randomize
              Players(LocPlayerN).Direction = Rnd() * 6.283185307
              Players(LocPlayerN).Posi.X = Rnd() * 800
              Players(LocPlayerN).Posi.Y = Rnd() * 600
              Call ShowMes("You where killed by " & Trim(Players(KPlayer).Name) & ".", CAction)
            Else
              Call ShowMes(Trim(Players(CPlayer).Name) & " was killed by " & Trim(Players(KPlayer).Name) & ".", CAction)
            End If
          End If
        Case KillXplo
          If CPlayer = LocPlayerN Then
            LocPSpeed.X = 0
            LocPSpeed.Y = 0
            Randomize
            Players(LocPlayerN).Direction = Rnd() * 6.283185307
            Players(LocPlayerN).Posi.X = Rnd() * 800
            Players(LocPlayerN).Posi.Y = Rnd() * 600
            Call ShowMes("You where killed in " & Trim(Players(KPlayer).Name) & "'s explosion.", CAction)
          ElseIf KPlayer = LocPlayerN Then
            Call ShowMes(Trim(Players(CPlayer).Name) & " where killed by your explosion.", CAction)
          Else
            Call ShowMes(Trim(Players(CPlayer).Name) & " was killed in " & Trim(Players(KPlayer).Name) & "'s explosion.", CAction)
          End If
        Case KillColl
          
        End Select
  End Select
End Sub

Private Sub SockU_DataArrival(ByVal bytesTotal As Long)
  Dim Data As String, P As Integer, CPlayer As Integer
  SockU.GetData Data
  For P = 0 To Len(Data) / 7 - 1
    CPlayer = Asc(Mid(Data, P * 7 + 1, 1))
    Players(CPlayer).Posi.X = (StrToInt(Mid(Data, P * 7 + 2, 2)) + 2 ^ 15) / 80
    Players(CPlayer).Posi.Y = (StrToInt(Mid(Data, P * 7 + 4, 2)) + 2 ^ 15) / 80
    If StrToInt(Mid(Data, P * 7 + 6, 2)) < 0 Then
      Players(CPlayer).Engine = True
      Players(CPlayer).Direction = (StrToInt(Mid(Data, P * 7 + 6, 2)) + 2 ^ 15) / 5000
    Else
      Players(CPlayer).Engine = False
      Players(CPlayer).Direction = StrToInt(Mid(Data, P * 7 + 6, 2)) / 5000
    End If
  Next P
End Sub

Public Sub SockU_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  SockU.Close
  SockU.Listen
  Call ShowMes("The UdP socket caused an error (" & Description & "). Connection closed.")
  Disconnect
End Sub

Public Sub SockT_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  SockU.Close
  SockU.Listen
  Call ShowMes("The TCP socket caused an error (" & Description & "). Connection closed.")
  Disconnect
End Sub

Public Sub ConnectSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  SockU.Close
  SockU.Listen
  Call ShowMes("The connection socket caused an error (" & Description & "). Connection closed.")
  Disconnect
End Sub
