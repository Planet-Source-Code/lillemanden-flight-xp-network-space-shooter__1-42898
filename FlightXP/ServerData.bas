Attribute VB_Name = "ServerData"
Option Explicit

Public Const GameSizeX As Integer = 800
Public Const GameSizeY As Integer = 600
Public Const MaxShots As Integer = 300
Public Const MinTickCount As Integer = 10
Public Const ColDecPrec As Single = 0.001
Public Const DeadTime As Single = 5
Public Const XploDeadly As Single = 0.5
Public Const XploRadius As Single = 25

Public Longside As Double

Public Codecolors(4) As Long
Public Const CServer As Integer = 0
Public Const CSSpeak As Integer = 1
Public Const CPSpeak As Integer = 2
Public Const CAction As Integer = 3
Public Const CPubMsg As Integer = 4


Public Running As Boolean

Public NPlayers As Integer
Public Players() As Player
Public NShots() As Integer
Public Shots() As Shot

Public Type Pos
  X As Single
  Y As Single
End Type

Public Type Shot
  Speed As Pos
  Posi As Pos
  InGame As Boolean
End Type

Public Type Player
  Name As String * 25
  Color As Long
  Posi As Pos
  Direction As Single
  Engine As Boolean
  InGame As Boolean
  Joining As Boolean
  Xploing As Single
End Type

Public Type Settings
  Maxplayers As Integer
  Port As Long
End Type

Public Const ConFull As String * 1 = 0
Public Const ConRoom As String * 1 = 1

Public Const TJoinInfo As String * 1 = 0
Public Const TJoin As String * 1 = 1
Public Const TShot As String * 1 = 2
Public Const TSay As String * 1 = 3
Public Const TChangeInfo As String * 1 = 4
Public Const TQuit As String * 1 = 5
Public Const TKick As String * 1 = 6
Public Const TKill As String * 1 = 7

Public Const KillShot As String * 1 = 1
Public Const KillXplo As String * 1 = 2
Public Const KillColl As String * 1 = 3

Public Function NAPlayers() As Integer
  Dim I As Integer
  NAPlayers = 0
  For I = 0 To NPlayers - 1
    If Players(I).InGame And Not (Players(I).Joining) Then NAPlayers = NAPlayers + 1
  Next I
End Function

Public Function NFreeSlots() As Integer
  Dim I As Integer
  NFreeSlots = 0
  For I = 0 To NPlayers - 1
    If Players(I).InGame Then NFreeSlots = NFreeSlots + 1
  Next I
End Function

Public Function FreePlayers() As Integer
  Dim I As Integer
  FreePlayers = -1
  For I = 0 To NPlayers - 1
    If Players(I).InGame = False Then
      FreePlayers = I
      Exit For
    End If
  Next I
End Function

Public Sub ResetPlayers()
  Dim I As Integer
  For I = 0 To NPlayers - 1
    Players(I).InGame = False
  Next I
End Sub

Public Sub SetCodeColors()
  Codecolors(0) = RGB(200, 20, 20)
  Codecolors(1) = RGB(0, 200, 0)
  Codecolors(2) = RGB(0, 0, 200)
  Codecolors(3) = RGB(50, 50, 50)
  Codecolors(4) = RGB(130, 20, 20)
End Sub

Public Sub ShowMes(Message As String, Optional Code As Integer = 0)
  ServerMain.tMessages.SelStart = Len(ServerMain.tMessages.Text)
  ServerMain.tMessages.SelColor = Codecolors(Code)
  ServerMain.tMessages.SelText = vbCrLf & Message
End Sub

Public Function ActiveShots(Player As Integer) As String
  ActiveShots = IntToStr(NShots(Player))
  Dim I As Integer
  For I = 0 To MaxShots - 1
    If Shots(Player, I).InGame Then
      ActiveShots = ActiveShots & IntToStr(I) & IntToStr(Shots(Player, I).Posi.X * 80 - 2 ^ 15) & IntToStr(Shots(Player, I).Posi.Y * 80 - 2 ^ 15) & IntToStr(Shots(Player, I).Speed.X * 30) & IntToStr(Shots(Player, I).Speed.Y * 30)
    End If
  Next I
End Function

Public Function IntToStr(I As Integer) As String
  IntToStr = Chr((I + 32768) And 255) & Chr(((I + 32768) And 65280) \ 256)
End Function

Public Function StrToInt(S As String) As Integer
  StrToInt = (Asc(S) + CLng(Asc(Right(S, 1))) * 256) - 32768
End Function

Public Sub PlayerJoins(Player As Integer)
  Dim P As Integer
  For P = 0 To NPlayers - 1
    If Players(P).InGame And P <> Player Then
      ServerMain.SockT(P).SendData TJoin & Chr(Player) & Players(Player).Name & Chr(Players(Player).Color And 255) & Chr((Players(Player).Color And 65280) \ 256) & Chr((Players(Player).Color And 16776960) \ 65536)
      DoEvents
    End If
  Next P
End Sub

Public Sub SendShot(Player As Integer, Shot As Integer)
  Dim P As Integer
  For P = 0 To NPlayers - 1
    If Players(P).InGame And P <> Player Then
      DoEvents
      ServerMain.SockT(P).SendData TShot & Chr(Player) & IntToStr(Shot) & IntToStr(Shots(Player, Shot).Posi.X * 80 - 2 ^ 15) & IntToStr(Shots(Player, Shot).Posi.Y * 80 - 2 ^ 15) & IntToStr(Shots(Player, Shot).Speed.X * 30) & IntToStr(Shots(Player, Shot).Speed.Y * 30)
      DoEvents
    End If
  Next P
End Sub

Public Sub SendQuit(Player As Integer)
  Dim P As Integer
  For P = 0 To NPlayers - 1
    If Players(P).InGame And P <> Player Then
      DoEvents
      ServerMain.SockT(P).SendData TQuit & Chr(Player)
      DoEvents
    End If
  Next P
End Sub

Public Sub SendChangeInfo(Player As Integer)
  Dim P As Integer
  For P = 0 To NPlayers - 1
    If Players(P).InGame And P <> Player Then
      DoEvents
      ServerMain.SockT(P).SendData TChangeInfo & Chr(Player) & Players(Player).Name & Chr(Players(Player).Color And 255) & Chr((Players(Player).Color And 65280) \ 256) & Chr((Players(Player).Color And 16776960) \ 65536)
      DoEvents
    End If
  Next P
End Sub

Public Sub SendSay(Player As Integer, Text As String)
  Dim P As Integer
  For P = 0 To NPlayers - 1
    If Players(P).InGame And P <> Player Then
      DoEvents
      ServerMain.SockT(P).SendData TSay & Chr(Player) & Text
      DoEvents
    End If
  Next P
End Sub

Public Sub SendKick(Player As Integer)
  Dim P As Integer
  For P = 0 To NPlayers - 1
    If Players(P).InGame Then
      DoEvents
      ServerMain.SockT(P).SendData TKick & Chr(Player)
      DoEvents
    End If
  Next P
  ServerMain.SockT(Player).Close
  ServerMain.SockU(Player).Close
  Players(Player).InGame = False
End Sub

Public Sub SendPPosi()
  On Error GoTo WinsockSucks
  Dim P As Integer, PosiData As String, I As Integer
  For P = 0 To NPlayers - 1
    If Players(P).InGame And Not (Players(P).Joining) Then
      PosiData = ""
      For I = 0 To NPlayers - 1
        If I <> P And Players(I).InGame And Not (Players(I).Joining) Then PosiData = PosiData & Chr(I) & IntToStr(Players(I).Posi.X * 80 - 2 ^ 15) & IntToStr(Players(I).Posi.Y * 80 - 2 ^ 15) & IntToStr(Players(I).Direction * 5000 + IIf(Players(I).Engine, -2 ^ 15, 0))
      Next I
      If PosiData <> "" Then
        DoEvents
        ServerMain.SockU(P).SendData PosiData
        DoEvents
      End If
    End If
  Next P
  Exit Sub
WinsockSucks:
  ShowMes ("Winsock caused an error becase it sucks (it will be ignored).")
End Sub

Public Sub SendKill(Player As Integer, Killer As Integer, Shot As Integer, KillType As Integer)
  Dim P As Integer
  For P = 0 To NPlayers - 1
    If Players(P).InGame And Not (Players(P).Joining) Then
      DoEvents
      ServerMain.SockT(P).SendData TKill & KillType & Chr(Player) & Chr(Killer) & IntToStr(Shot)
      DoEvents
    End If
  Next P
End Sub

Public Sub ServerCommand(Command As String)
  If LCase(Left(Command, 5)) = "kick " Then
    If IsNumeric(Mid(Command, 6)) Then
      If Mid(Command, 6) >= 0 And Mid(Command, 6) < NPlayers And CDbl(Mid(Command, 6)) = Int(Mid(Command, 6)) Then
        If Players(Mid(Command, 6)).InGame Then
          SendKick (Int(Mid(Command, 6)))
          Call ShowMes(Trim(Players(Mid(Command, 6)).Name) & " at playerslot " & Mid(Command, 6) & " has been kicked.", CPubMsg)
          ServerMain.lPlayers.Caption = "Players: " & NFreeSlots & "/" & NPlayers
        Else
          ShowMes ("'" & Mid(Command, 6) & "' is an empty playerslot!")
        End If
      Else
        ShowMes ("'" & Mid(Command, 6) & "' is not a valid playerslot!")
      End If
    Else
      ShowMes ("'" & Mid(Command, 6) & "' is not a number!")
    End If
  ElseIf LCase(Left(Command, 10)) = "playerlist" Then
    Dim P As Integer
    For P = 0 To NPlayers - 1
      If Players(P).InGame Then
        ShowMes ("Playerslot " & P & ": " & Trim(Players(P).Name))
      End If
    Next P
  ElseIf Command = "?" Then
    ShowMes ("\? - The help you are seeing now.")
    ShowMes ("\playerlist - Lists all the players currently connected.")
    ShowMes ("\kick # - Kick the player on playerslot #.")
  End If
End Sub

Public Sub ColDecShotPlayer(Player As Integer)
  Dim S As Integer, P As Integer, a As Double, b As Double, c As Double, Pa As Pos, Pb As Pos, Pc As Pos
  Pa.X = Players(Player).Posi.X - Sin(Players(Player).Direction) * 10
  Pa.Y = Players(Player).Posi.Y + Cos(Players(Player).Direction) * 10
  Pb.X = Players(Player).Posi.X + Cos(Players(Player).Direction) * 5 + Sin(Players(Player).Direction) * 5
  Pb.Y = Players(Player).Posi.Y + Sin(Players(Player).Direction) * 5 + Cos(Players(Player).Direction) * -5
  Pc.X = Players(Player).Posi.X + Cos(Players(Player).Direction) * -5 + Sin(Players(Player).Direction) * 5
  Pc.Y = Players(Player).Posi.Y + Sin(Players(Player).Direction) * -5 + Cos(Players(Player).Direction) * -5
  For P = 0 To NPlayers - 1
    If DeadTime - Players(P).Xploing < XploDeadly And Sqr((Players(P).Posi.X - Players(Player).Posi.X) ^ 2 + (Players(P).Posi.Y - Players(Player).Posi.Y) ^ 2) < XploRadius Then
      Call SendKill(Player, P, 0, KillXplo)
      Players(Player).Xploing = DeadTime
      Call ShowMes(Trim(Players(Player).Name) & " was killed in " & Trim(Players(P).Name) & "'s explosion.", CAction)
      Exit Sub
    End If
    For S = 0 To MaxShots - 1
      If Shots(P, S).InGame Then
        If Sqr((Shots(P, S).Posi.X - Players(Player).Posi.X) ^ 2 + (Shots(P, S).Posi.Y - Players(Player).Posi.Y) ^ 2) < 10 Then
          a = Sqr((Shots(P, S).Posi.X - Pa.X) ^ 2 + (Shots(P, S).Posi.Y - Pa.Y) ^ 2)
          b = Sqr((Shots(P, S).Posi.X - Pb.X) ^ 2 + (Shots(P, S).Posi.Y - Pb.Y) ^ 2)
          c = Sqr((Shots(P, S).Posi.X - Pc.X) ^ 2 + (Shots(P, S).Posi.Y - Pc.Y) ^ 2)
          If 6.283185307 - Abs(ArcCos((a ^ 2 + b ^ 2 - Longside) / (2 * a * b))) - Abs(ArcCos((a ^ 2 + c ^ 2 - Longside) / (2 * a * c))) - ArcCos((b ^ 2 + c ^ 2 - 100) / (2 * b * c)) < ColDecPrec Then
            Shots(P, S).InGame = False
            NShots(P) = NShots(P) - 1
            Call SendKill(Player, P, S, KillShot)
            Players(Player).Xploing = DeadTime
            If P = Player Then
              Call ShowMes(Trim(Players(Player).Name) & " committed suicide!", CAction)
            Else
              Call ShowMes(Trim(Players(Player).Name) & " was killed by " & Trim(Players(P).Name) & ".", CAction)
            End If
            Exit Sub
          End If
        End If
      End If
    Next S
  Next P
End Sub

Public Function ArcCos(X As Double) As Double
  If X >= 1 Or X <= -1 Then
    ArcCos = 0
  Else
    ArcCos = 2 * Atn(1) - Atn(X / Sqr(1 - X * X))
  End If
End Function
