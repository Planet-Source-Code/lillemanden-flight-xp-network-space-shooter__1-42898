VERSION 5.00
Begin VB.Form Gameoptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Game Options"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Engine Effekt"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.CheckBox SimpleEngine 
         Caption         =   "Use simple effekt"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.HScrollBar ParticScroll 
         Height          =   255
         LargeChange     =   200
         Left            =   240
         Max             =   1000
         Min             =   100
         SmallChange     =   50
         TabIndex        =   2
         Top             =   600
         Value           =   300
         Width           =   3495
      End
      Begin VB.Label Npartic 
         Caption         =   "300"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Particles per engine:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Update 
      Caption         =   "Update"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   3975
   End
End
Attribute VB_Name = "Gameoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  If NThrustFire <> 0 Then
    ParticScroll.Value = NThrustFire
    Npartic = ParticScroll.Value
  Else
    ParticScroll.Enabled = False
    Npartic = "N/A"
    SimpleEngine.Value = 1
  End If
End Sub

Private Sub ParticScroll_Change()
  Npartic = ParticScroll.Value
End Sub

Private Sub SimpleEngine_Click()
  If SimpleEngine.Value = 0 Then
    ParticScroll.Enabled = True
    Npartic = ParticScroll.Value
  Else
    ParticScroll.Enabled = False
    Npartic = "N/A"
  End If
End Sub

Private Sub Update_Click()
  NThrustFire = 0
  If SimpleEngine.Value = 0 Then
    If Playing Then
      Dim WaitTime As Long
      WaitTime = GetTickCount
      While GetTickCount - WaitTime < 0.5 And GetTickCount - WaitTime > 0
        DoEvents
      Wend
      ReDim ThrustFire(NPlayers - 1, ParticScroll.Value - 1)
    End If
    NThrustFire = ParticScroll.Value
  End If
  Gameoptions.Hide
End Sub
