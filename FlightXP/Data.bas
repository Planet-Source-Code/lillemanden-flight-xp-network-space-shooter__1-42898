Attribute VB_Name = "Data"
Option Explicit

Public Const GameSizeX As Integer = 800
Public Const GameSizeY As Integer = 600
Public Const MaxShots As Integer = 300

Public Const TurnSpeed As Integer = 3.5
Public Const ThrustSpeed As Integer = 400
Public Const StopSpeed As Single = 0.5
Public Const ShotSpeed As Integer = 350
Public Const ShootSpeed As Integer = 150
Public Const SpamTime As Integer = 250
Public Const MinTickCount As Integer = 10
Public Const XploAccl As Single = -380
Public Const DeadTime As Long = 5

Public Playing As Boolean
Public HasFocus As Boolean

Public LastSay As Long
Public LastShot As Long

Public NPlayers As Integer
Public Players() As Player
Public NShots() As Integer
Public Shots() As Shot
Public Xplos() As Explo
Public ThrustFire() As Thrust
Public NThrustFire As Integer

Public LocPSpeed As Pos
Public LocPlayerN As Integer
Public LocPName As String * 25
Public LocPColor As Long

Public Type Pos
  X As Single
  Y As Single
End Type

Public Type Thrust
  Color As Long
  Dies As Single
  Posi As Pos
  Speed As Pos
End Type

Public Type Shot
  Speed As Pos
  Posi As Pos
  Ingame As Boolean
End Type

Public Type Player
  Name As String * 25
  Color As Long
  Posi As Pos
  Direction As Single
  Engine As Boolean
  Ingame As Boolean
  Xploing As Single
End Type

Public Type Explo
  Posi As Pos
  Size As Single
  Speed As Single
  Ingame As Boolean
End Type

Public Type Settings
  Name As String * 25
  Color As Long
  JoinHost As String * 50
  JoinPort As Long
  EngineEffect As Integer
End Type

Public Codecolors(4) As Long
Public Const CSSpeak As Integer = 0
Public Const CPSpeak As Integer = 1
Public Const CAction As Integer = 2
Public Const CPubMsg As Integer = 3
Public Const CLocMsg As Integer = 4


Public Sub CreateShot(Posi As Pos, Speed As Pos)
  If Posi.X < 0 Or Posi.X > GameSizeX Or Posi.Y < 0 Or Posi.Y > GameSizeY Then Exit Sub
  Dim I As Integer
  For I = 0 To MaxShots
    If Shots(LocPlayerN, I).Ingame = False Then
      Shots(LocPlayerN, I).Ingame = True
      Shots(LocPlayerN, I).Posi = Posi
      Shots(LocPlayerN, I).Speed = Speed
      Exit For
    End If
  Next I
  NShots(LocPlayerN) = NShots(LocPlayerN) + 1
  Main.SockT.SendData TShot & IntToStr(I) & IntToStr(Shots(LocPlayerN, I).Posi.X * 80 - 2 ^ 15) & IntToStr(Shots(LocPlayerN, I).Posi.Y * 80 - 2 ^ 15) & IntToStr(Shots(LocPlayerN, I).Speed.X * 30) & IntToStr(Shots(LocPlayerN, I).Speed.Y * 30)
  DoEvents
End Sub

Public Sub ShowMes(Message As String, Optional Code As Integer = 4)
  fMessage.tMessage.SelStart = Len(fMessage.tMessage.Text)
  fMessage.tMessage.SelColor = Codecolors(Code)
  fMessage.tMessage.SelText = vbCrLf & Message
End Sub

Public Sub SetCodeColors()
  Codecolors(0) = RGB(0, 200, 0)
  Codecolors(1) = RGB(0, 0, 200)
  Codecolors(2) = RGB(50, 50, 50)
  Codecolors(3) = RGB(130, 20, 20)
  Codecolors(4) = RGB(200, 50, 50)
End Sub

Public Function IntToStr(I As Integer) As String
  IntToStr = Chr((I + 32768) And 255) & Chr(((I + 32768) And 65280) \ 256)
End Function

Public Function StrToInt(S As String) As Integer
  StrToInt = (Asc(S) + CLng(Asc(Right(S, 1))) * 256) - 32768
End Function

Public Function LocInfo() As String
   LocInfo = LocPName & Chr(LocPColor And 255) & Chr((LocPColor And 65280) \ 256) & Chr((LocPColor And 16776960) \ 65536)
End Function
