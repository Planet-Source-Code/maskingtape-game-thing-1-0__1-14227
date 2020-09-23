VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Game!"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2625
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "GO!"
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Skill Level"
      Height          =   1335
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
      Begin VB.OptionButton optHard 
         Caption         =   "I'm God!"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optMedium 
         Caption         =   "I'm pretty good."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optEasy 
         Caption         =   "I suck!"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Game Speed"
      Height          =   1335
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1935
      Begin VB.OptionButton optWarp10 
         Caption         =   "Warp 10!"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optNotss 
         Caption         =   "Not so slow."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optSlow 
         Caption         =   "Slow."
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.TextBox txtGT 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "20"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "How many squares you must clear."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If optEasy.Value = True Then Form1.Line5.Y1 = 2400: Form1.Line5.Y2 = 2400: Level = 1
If optMedium.Value = True Then Form1.Line5.Y1 = 2520: Form1.Line5.Y2 = 2520: Level = 2
If optHard.Value = True Then Form1.Line5.Y1 = 2640: Form1.Line5.Y2 = 2640: Level = 3

Counter = txtGT.Text
Custom = True

Form1.lblLevel.Caption = "---"
Form1.lblStage.Caption = "---"
Form1.mnuCustom.Caption = "Stop"
Form1.mnuArcade.Enabled = False
    
Call ShowSquares
Call Start
Unload Me
End Sub

Private Sub Form_Load()
Speed(1) = 30
Speed(2) = 40
Speed(3) = 25
Speed(4) = 10
Speed(5) = 35
Speed(6) = 20
End Sub

Private Sub optNotss_Click()
Speed(1) = 40
Speed(2) = 30
Speed(3) = 50
Speed(4) = 45
Speed(5) = 55
Speed(6) = 60
End Sub

Private Sub optSlow_Click()
Speed(1) = 30
Speed(2) = 40
Speed(3) = 25
Speed(4) = 10
Speed(5) = 35
Speed(6) = 20
End Sub

Private Sub optWarp10_Click()
Speed(1) = 65
Speed(2) = 75
Speed(3) = 60
Speed(4) = 45
Speed(5) = 70
Speed(6) = 65
End Sub
