Attribute VB_Name = "Net"
Option Explicit

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

Public Sub SendLocPlayer()
  Main.SockU.SendData IntToStr(Players(LocPlayerN).Posi.X * 80 - 2 ^ 15) & IntToStr(Players(LocPlayerN).Posi.Y * 80 - 2 ^ 15) & IntToStr(Players(LocPlayerN).Direction * 5000 + IIf(Players(LocPlayerN).Engine, -2 ^ 15, 0))
End Sub
