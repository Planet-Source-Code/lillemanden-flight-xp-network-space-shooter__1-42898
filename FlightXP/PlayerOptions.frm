VERSION 5.00
Begin VB.Form PlayerOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Player Options"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar ChangeColor 
      Height          =   255
      Left            =   2280
      Max             =   10
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.PictureBox ShowColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      ScaleHeight     =   225
      ScaleWidth      =   1665
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox PName 
      Height          =   285
      Left            =   600
      MaxLength       =   25
      TabIndex        =   3
      Text            =   "Player"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Update 
      Caption         =   "Update"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "PlayerOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NColors As Integer = 20
Private Colors(NColors) As Long

Private Sub ChangeColor_Change()
  ShowColor.BackColor = Colors(ChangeColor.Value)
End Sub

Private Sub Form_Load()
  ChangeColor.Max = NColors
  Dim I As Integer
  For I = 0 To NColors
    Colors(I) = ((2 ^ 24 - 75) / (NColors + 1)) * I + 75
  Next I
  ShowColor.BackColor = LocPColor
  PName.Text = Trim(LocPName)
End Sub

Private Sub Update_Click()
  If PName = "" Then
    Call MsgBox("You didn't fill in a name!", vbCritical, "No Name!")
  Else
    LocPColor = ShowColor.BackColor
    LocPName = PName
    If Playing Then
      Call ShowMes("You changed playeroptions" & IIf(Players(LocPlayerN).Name <> LocPName, ", and is now known as " & LocPName & ".", "."), CPubMsg)
      Players(LocPlayerN).Color = LocPColor
      Players(LocPlayerN).Name = LocPName
      Main.SockT.SendData TChangeInfo & LocInfo
    End If
    PlayerOptions.Hide
  End If
End Sub
