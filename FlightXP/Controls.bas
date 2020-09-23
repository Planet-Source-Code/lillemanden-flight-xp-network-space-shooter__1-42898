Attribute VB_Name = "Controls"
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Keys(4) As Boolean


Public Sub UControl()
  If HasFocus Then
    Keys(0) = IIf(GetAsyncKeyState(vbKeyUp) <> 0, True, False)
    Keys(1) = IIf(GetAsyncKeyState(vbKeyLeft) <> 0, True, False)
    Keys(2) = IIf(GetAsyncKeyState(vbKeyRight) <> 0, True, False)
    Keys(3) = IIf(GetAsyncKeyState(vbKeySpace) <> 0, True, False)
    Keys(4) = IIf(GetAsyncKeyState(vbKeyReturn) <> 0, True, False)
  End If
  If Players(LocPlayerN).Xploing = 0 Then
    If Keys(1) Then
      Players(LocPlayerN).Direction = Players(LocPlayerN).Direction - TurnSpeed * Rate
      If Players(LocPlayerN).Direction < 0 Then Players(LocPlayerN).Direction = 6.283185307 - Players(LocPlayerN).Direction
    End If
    If Keys(2) Then
      Players(LocPlayerN).Direction = Players(LocPlayerN).Direction + TurnSpeed * Rate
      If Players(LocPlayerN).Direction > 6.283185307 Then Players(LocPlayerN).Direction = Players(LocPlayerN).Direction - 6.283185307
    End If
    If Keys(0) Then
      LocPSpeed.X = LocPSpeed.X - ThrustSpeed * Rate * Sin(Players(LocPlayerN).Direction)
      LocPSpeed.Y = LocPSpeed.Y + ThrustSpeed * Rate * Cos(Players(LocPlayerN).Direction)
      Players(LocPlayerN).Engine = True
    Else
      Players(LocPlayerN).Engine = False
    End If
    If LocPSpeed.X <> 0 Then LocPSpeed.X = LocPSpeed.X - StopSpeed * Rate * (LocPSpeed.X / (Abs(LocPSpeed.X) + Abs(LocPSpeed.Y))) * (Abs(LocPSpeed.X) + Abs(LocPSpeed.Y))
    If LocPSpeed.Y <> 0 Then LocPSpeed.Y = LocPSpeed.Y - StopSpeed * Rate * (LocPSpeed.Y / (Abs(LocPSpeed.X) + Abs(LocPSpeed.Y))) * (Abs(LocPSpeed.X) + Abs(LocPSpeed.Y))
    Players(LocPlayerN).Posi.X = Players(LocPlayerN).Posi.X + LocPSpeed.X * Rate
    Players(LocPlayerN).Posi.Y = Players(LocPlayerN).Posi.Y + LocPSpeed.Y * Rate
    If Players(LocPlayerN).Posi.X < 0 Then Players(LocPlayerN).Posi.X = Players(LocPlayerN).Posi.X + GameSizeX
    If Players(LocPlayerN).Posi.X > GameSizeX Then Players(LocPlayerN).Posi.X = Players(LocPlayerN).Posi.X - GameSizeX
    If Players(LocPlayerN).Posi.Y < 0 Then Players(LocPlayerN).Posi.Y = Players(LocPlayerN).Posi.Y + GameSizeY
    If Players(LocPlayerN).Posi.Y > GameSizeY Then Players(LocPlayerN).Posi.Y = Players(LocPlayerN).Posi.Y - GameSizeY
    If Keys(3) And GetTickCount - LastShot > ShootSpeed Then
      LastShot = GetTickCount
      Dim SPosi As Pos
      Dim SSpeed As Pos
      SPosi.X = Sin(-Players(LocPlayerN).Direction) * 12 + Players(LocPlayerN).Posi.X
      SPosi.Y = Cos(-Players(LocPlayerN).Direction) * 12 + Players(LocPlayerN).Posi.Y
      SSpeed.X = Sin(-Players(LocPlayerN).Direction) * ShotSpeed + LocPSpeed.X
      SSpeed.Y = Cos(-Players(LocPlayerN).Direction) * ShotSpeed + LocPSpeed.Y
      Call CreateShot(SPosi, SSpeed)
    End If
  End If
  If Keys(4) And GetTickCount - LastSay > SpamTime Then
    Keys(4) = False
    fMessage.Show
    fMessage.SayText.Enabled = True
  End If
End Sub
