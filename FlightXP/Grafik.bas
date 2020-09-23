Attribute VB_Name = "Grafik"
Option Explicit
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Const ThrustMaxT As Single = 0.15
Private Const ThrustMinT As Single = 0.05
Private Const ThrustMaxS As Single = 100
Private Const ThrustMinS As Single = 50
Private Const ThrustRndS As Single = 1.5
Private Const ThrustRate As Single = 0.1

Public Sub DrawPlayer(PlayerN As Integer)
  Dim C As Long
  Dim P As Pos
  Dim D As Single
  Dim Si As Single
  Dim Co As Single
  C = Players(PlayerN).Color
  P = Players(PlayerN).Posi
  D = Players(PlayerN).Direction
  Si = Sin(D)
  Co = Cos(D)
  Main.Game.Line (-Si * 10 + P.x, Co * 10 + P.y)-(Co * 5 - Si * -5 + P.x, Si * 5 + Co * -5 + P.y), C
  Main.Game.Line (-Si * 10 + P.x, Co * 10 + P.y)-(Co * -5 - Si * -5 + P.x, Si * -5 + Co * -5 + P.y), C
  Main.Game.Line (Co * 5 - Si * -5 + P.x, Si * 5 + Co * -5 + P.y)-(Co * -5 - Si * -5 + P.x, Si * -5 + Co * -5 + P.y), C
  Main.Game.Line (-Si * 10 + P.x, Co * 10 + P.y)-(Co * 2.5 - Si * -5 + P.x, Si * 2.5 + Co * -5 + P.y), C
  Main.Game.Line (-Si * 10 + P.x, Co * 10 + P.y)-(Co * -2.5 - Si * -5 + P.x, Si * -2.5 + Co * -5 + P.y), C
  Main.Game.Line (-Si * 10 + P.x, Co * 10 + P.y)-(-Si * -5 + P.x, Co * -5 + P.y), C
  If NThrustFire <> 0 Then
    Dim T As Integer, MaxOut As Integer, RVal As Single
    Randomize
    MaxOut = Rnd() * (NThrustFire / ThrustRate) * Rate
    For T = 0 To NThrustFire - 1
      If Players(PlayerN).Engine And ThrustFire(PlayerN, T).Dies <= 0 And MaxOut > 0 Then
        MaxOut = MaxOut - 1
        ThrustFire(PlayerN, T).Dies = Rnd() * ThrustMaxT + ThrustMinT
        ThrustFire(PlayerN, T).Color = RGB(Rnd() * 175 + 75, Rnd() * 75 + 25, Rnd() * 75)
        RVal = Rnd() * 8 - 4
        ThrustFire(PlayerN, T).Posi.x = Co * RVal + Si * 5 + P.x
        ThrustFire(PlayerN, T).Posi.y = Si * RVal - Co * 5 + P.y
        RVal = Rnd() * ThrustMaxS + ThrustMinS
        ThrustFire(PlayerN, T).Speed.x = Si * RVal * Rnd() * ThrustRndS
        ThrustFire(PlayerN, T).Speed.y = -Co * RVal * Rnd() * ThrustRndS
      End If
      If ThrustFire(PlayerN, T).Dies > 0 Then
        Call SetPixel(Main.Game.hdc, ThrustFire(PlayerN, T).Posi.x, ThrustFire(PlayerN, T).Posi.y, ThrustFire(PlayerN, T).Color)
        ThrustFire(PlayerN, T).Posi.x = ThrustFire(PlayerN, T).Posi.x + ThrustFire(PlayerN, T).Speed.x * Rate
        ThrustFire(PlayerN, T).Posi.y = ThrustFire(PlayerN, T).Posi.y + ThrustFire(PlayerN, T).Speed.y * Rate
        ThrustFire(PlayerN, T).Dies = ThrustFire(PlayerN, T).Dies - Rate
      End If
    Next T
  Else
    If Players(PlayerN).Engine Then
      Randomize
      Dim EColor As Long
      EColor = RGB(Rnd() * 100 + 100, Rnd() * 100, 0)
      Main.Game.Line (Co * 4 - Si * -6 + P.x, Si * 4 + Co * -6 + P.y)-(Co * -4 - Si * -6 + P.x, Si * -4 + Co * -6 + P.y), EColor
      Main.Game.Line (Co * 3 - Si * -8 + P.x, Si * 3 + Co * -8 + P.y)-(Co * -3 - Si * -8 + P.x, Si * -3 + Co * -8 + P.y), EColor
      Main.Game.Line (Co * 2 - Si * -10 + P.x, Si * 2 + Co * -10 + P.y)-(Co * -2 - Si * -10 + P.x, Si * -2 + Co * -10 + P.y), EColor
    End If
  End If
End Sub

Public Sub DrawShot(PlayerN As Integer, ShotN As Integer)
  Call SetPixel(Main.Game.hdc, Shots(PlayerN, ShotN).Posi.x, Shots(PlayerN, ShotN).Posi.y, RGB(0, 255, 0))
End Sub

Public Sub SetXplo(XploN As Integer, Optional x As Single = -10, Optional y As Single = -10, Optional StartSize As Single = 5, Optional Speed As Single = 120)
  If y = -10 Then
    x = Players(XploN).Posi.x
    y = Players(XploN).Posi.y
  End If
  Xplos(XploN).Posi.x = x
  Xplos(XploN).Posi.y = y
  Xplos(XploN).Size = StartSize
  Xplos(XploN).Speed = Speed
  Xplos(XploN).Ingame = True
End Sub

Public Sub DoXplos()
  Dim x As Integer, S As Integer, L As Single
  Randomize
  For x = 0 To NPlayers - 1
    If Xplos(x).Ingame Then
      For S = 1 To Xplos(x).Size
        Main.Game.Circle (Xplos(x).Posi.x, Xplos(x).Posi.y), S, RGB(Rnd() * 125 + 131, Rnd() * 75, Rnd() * 25)
      Next S
      For S = 1 To Xplos(x).Size * 4
        L = Rnd() * 6.283185307
        Main.Game.Line (Xplos(x).Posi.x, Xplos(x).Posi.y)-(Cos(L) * Xplos(x).Size ^ 1.3 * Rnd() + Xplos(x).Posi.x, Sin(L) * Xplos(x).Size ^ 1.3 * Rnd() + Xplos(x).Posi.y), RGB(Rnd() * 100 + 100, Rnd() * 50 + 25, Rnd() * 50)
      Next S
      Xplos(x).Size = Xplos(x).Size + Xplos(x).Speed * Rate
      Xplos(x).Speed = Xplos(x).Speed + XploAccl * Rate
      If Xplos(x).Size < 0 Then Xplos(x).Ingame = False
    End If
  Next x
End Sub
