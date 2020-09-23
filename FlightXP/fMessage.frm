VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form fMessage 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Flight XP Messagewindow"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox SayText 
      Height          =   285
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3360
      Width           =   3255
   End
   Begin RichTextLib.RichTextBox tMessage 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5953
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"fMessage.frx":0000
   End
End
Attribute VB_Name = "fMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
  If fMessage.Width < 3375 Then fMessage.Width = 3375
  If fMessage.Height < 2610 Then fMessage.Height = 2610
  tMessage.Width = fMessage.Width - 120
  tMessage.Height = fMessage.Height - 675
  SayText.Width = fMessage.Width - 120
  SayText.Top = fMessage.Height - 665
End Sub

Private Sub Form_Activate()
  SayText.SetFocus
End Sub

Private Sub SayText_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SayText = Replace(SayText, vbCrLf, "")
    If SayText <> "" Then
      Main.SockT.SendData TSay & SayText
      Call ShowMes(Trim(LocPName) & ": " & SayText, CPSpeak)
      SayText = ""
    End If
    LastSay = GetTickCount
    Main.Game.SetFocus
  End If
End Sub
