Attribute VB_Name = "MainLoop"
Option Explicit
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Rate As Single

Public Sub GLoop()
  Dim OldTimer As Long, I As Integer, O As Integer, P As Integer
  LocPSpeed.X = 0
  LocPSpeed.Y = 0
  Randomize
  Players(LocPlayerN).Direction = Rnd() * 6.283185307
  Players(LocPlayerN).Posi.X = Rnd() * 800
  Players(LocPlayerN).Posi.Y = Rnd() * 600
  OldTimer = GetTickCount
  Main.Game.SetFocus
  LastShot = GetTickCount
  LastSay = GetTickCount
  While Playing
    If GetTickCount - OldTimer < 0 Then
      OldTimer = GetTickCount
    End If
    If GetTickCount - OldTimer > MinTickCount Then
      Rate = (GetTickCount - OldTimer) / 1000
      OldTimer = GetTickCount
      UControl
      If Players(LocPlayerN).Xploing = 0 Then SendLocPlayer
      Main.Game.Cls
      For I = 0 To NPlayers - 1
        If Players(I).Ingame Then
          If Players(I).Xploing = 0 Then
            DrawPlayer (I)
          Else
            Players(I).Xploing = Players(I).Xploing - Rate
            If Players(I).Xploing < 0 Then Players(I).Xploing = 0
          End If
        End If
      Next I
      For P = 0 To NPlayers - 1
        If NShots(P) <> 0 Then
          I = NShots(P)
          For O = 0 To MaxShots - 1
            If Shots(P, O).Ingame Then
              Call DrawShot(P, O)
              Shots(P, O).Posi.X = Shots(P, O).Posi.X + Shots(P, O).Speed.X * Rate
              Shots(P, O).Posi.Y = Shots(P, O).Posi.Y + Shots(P, O).Speed.Y * Rate
              If Shots(P, O).Posi.X < 0 Or Shots(P, O).Posi.X > GameSizeX Or Shots(P, O).Posi.Y < 0 Or Shots(P, O).Posi.Y > GameSizeY Then
                Shots(P, O).Ingame = False
                NShots(P) = NShots(P) - 1
              End If
              I = I - 1
              If I = 0 Then Exit For
            End If
          Next O
        End If
      Next P
      DoXplos
    End If
    DoEvents
  Wend
End Sub


