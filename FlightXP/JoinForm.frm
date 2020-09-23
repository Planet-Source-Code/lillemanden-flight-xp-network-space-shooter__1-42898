VERSION 5.00
Begin VB.Form JoinForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Join"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   1830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Join 
      Caption         =   "Join"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox JoinPort 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Text            =   "8500"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox JoinIP 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Text            =   "127.1.1.1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "JoinForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Join_Click()
  fMessage.Show
  fMessage.SayText.Enabled = True
  Playing = True
  Main.mJoin.Caption = "Disconnect"
  Main.mMessage.Enabled = True
  Main.mScore.Enabled = True
  Main.ConnectSock.Close
  Main.ConnectSock.RemoteHost = JoinIP
  Main.ConnectSock.RemotePort = JoinPort
  Main.ConnectSock.Connect
  JoinForm.Hide
  ShowMes ("Connecting to " & Trim(JoinIP) & ":" & JoinPort & ".")
End Sub
