Attribute VB_Name = "ServerMainLoop"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Rate As Single

Public Sub SLoop()
  Dim OldTimer As Long, P As Integer, O As Integer, I As Integer
  Running = True
  ReDim Players(NPlayers - 1)
  ReDim NShots(NPlayers - 1)
  ReDim Shots(NPlayers - 1, MaxShots - 1)
  ResetPlayers
  OldTimer = GetTickCount
  ShowMes ("Server is running.")
  While Running
    If GetTickCount - OldTimer < 0 Then
      OldTimer = GetTickCount
    End If
    If GetTickCount - OldTimer > MinTickCount Then
      Rate = (GetTickCount - OldTimer) / 1000
      OldTimer = GetTickCount
      For P = 0 To NPlayers - 1
        If NShots(P) <> 0 Then
          I = NShots(P)
          For O = 0 To MaxShots - 1
            If Shots(P, O).InGame Then
              Shots(P, O).Posi.X = Shots(P, O).Posi.X + Shots(P, O).Speed.X * Rate
              Shots(P, O).Posi.Y = Shots(P, O).Posi.Y + Shots(P, O).Speed.Y * Rate
              If Shots(P, O).Posi.X < 0 Or Shots(P, O).Posi.X > GameSizeX Or Shots(P, O).Posi.Y < 0 Or Shots(P, O).Posi.Y > GameSizeY Then
                Shots(P, O).InGame = False
                NShots(P) = NShots(P) - 1
              End If
              I = I - 1
              If I = 0 Then Exit For
            End If
          Next O
        End If
      Next P
      For P = 0 To NPlayers - 1
        If Players(P).InGame Then
          If Players(P).Xploing = 0 Then
            '<-------------------------player-player col detection
            ColDecShotPlayer (P)
          Else
            Players(P).Xploing = Players(P).Xploing - Rate
            If Players(P).Xploing < 0 Then Players(P).Xploing = 0
          End If
        End If
      Next P
      SendPPosi
    End If
    DoEvents
  Wend
End Sub
