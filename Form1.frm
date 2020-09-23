VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Thing 1.0"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4515
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      ScaleHeight     =   225
      ScaleWidth      =   1665
      TabIndex        =   26
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   3360
      TabIndex        =   22
      Top             =   840
      Width           =   1095
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Demo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Misses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   3360
      TabIndex        =   18
      Top             =   2400
      Width           =   1095
      Begin VB.Label lblMiss2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblMiss3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblMiss1 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Squares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   3360
      TabIndex        =   16
      Top             =   0
      Width           =   1095
      Begin VB.Label lblCounter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Demo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Timer Timer7 
      Interval        =   100
      Left            =   4560
      Top             =   1560
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5520
      Top             =   840
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5040
      Top             =   840
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4560
      Top             =   840
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5520
      Top             =   120
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3000
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2520
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2040
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1080
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5040
      Top             =   120
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4560
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblCongrats 
      Alignment       =   2  'Center
      Caption         =   "Congratulations!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   28
      Top             =   240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblGO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   27
      Top             =   1200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblStage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Demo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3360
      TabIndex        =   25
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Demo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3360
      TabIndex        =   24
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblGameThing 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Thing              1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   240
      TabIndex        =   15
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label lblPerfect 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Perfect!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   720
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblLevelO 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Level Over"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00800000&
      X1              =   240
      X2              =   3120
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblPaused 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PAUSED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblJ 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   3000
      Width           =   255
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00800000&
      X1              =   3120
      X2              =   3120
      Y1              =   0
      Y2              =   2880
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00800000&
      X1              =   2640
      X2              =   2640
      Y1              =   0
      Y2              =   2880
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00800000&
      X1              =   2160
      X2              =   2160
      Y1              =   0
      Y2              =   2880
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000C0&
      X1              =   120
      X2              =   3240
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   3000
      Width           =   255
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      X1              =   1200
      X2              =   1200
      Y1              =   0
      Y2              =   2880
   End
   Begin VB.Label lblS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   255
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      X1              =   3120
      X2              =   240
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   720
      X2              =   720
      Y1              =   2880
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   2880
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuArcade 
         Caption         =   "Arcade Mode"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuCustom 
         Caption         =   "Custom Mode"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sound As Long

If KeyCode = vbKeyA Then
    lblA.ForeColor = vbRed
    If Paused = True Or Started = False Then
    Else
    If Picture1.Top >= 2040 And Picture1.Top <= 2570 And Level = 1 Then
        Picture1.Top = 120
        Score = Score + 10
        Call CheckCounter
    ElseIf Picture1.Top >= 2160 And Picture1.Top <= 2570 And Level = 2 Then
        If Hard(1) = True Then
            If Press(1) = 1 Then
                Picture1.Top = 120
                Score = Score + 20
                Picture1.BackColor = vbWhite
                Press(1) = 0
                Hard(1) = False
                Call CheckCounter
            Else
                Let Press(1) = Press(1) + 1
            End If
        Else
            Picture1.Top = 120
            Score = Score + 10
            Call CheckCounter
            Call Block1Hard
        End If
    ElseIf Picture1.Top >= 2280 And Picture1.Top <= 2570 And Level = 3 Then
        If Hard(1) = True Then
            If Press(1) = 1 Then
                Picture1.Top = 120
                Score = Score + 20
                Picture1.BackColor = vbWhite
                Press(1) = 0
                Hard(1) = False
                Call CheckCounter
            Else
                Let Press(1) = Press(1) + 1
            End If
        Else
            Picture1.Top = 120
            Score = Score + 10
            Call CheckCounter
            Call Block1Hard
        End If
    Else
        sound = sndPlaySound(App.Path + "\buzz.wav", 1)
        Score = Score - 10
        Fdup = True
        If Score < -150 Then
            Call HideSquares
            Call StopG
            lblGO.Visible = True
        End If
    End If
    End If
End If

If KeyCode = vbKeyS Then
    lblS.ForeColor = vbRed
    If Paused = True Or Started = False Then
    Else
    If Picture2.Top >= 2040 And Picture2.Top <= 2570 And Level = 1 Then
        Picture2.Top = 120
        Score = Score + 10
        Call CheckCounter
    ElseIf Picture2.Top >= 2160 And Picture2.Top <= 2570 And Level = 2 Then
        If Hard(2) = True Then
            If Press(2) = 1 Then
                Picture2.Top = 120
                Score = Score + 20
                Picture2.BackColor = vbWhite
                Press(2) = 0
                Hard(2) = False
                Call CheckCounter
            Else
                Let Press(2) = Press(2) + 1
            End If
        Else
            Picture2.Top = 120
            Score = Score + 10
            Call CheckCounter
            Call Block2Hard
        End If
    ElseIf Picture2.Top >= 2280 And Picture2.Top <= 2570 And Level = 3 Then
        If Hard(2) = True Then
            If Press(2) = 1 Then
                Picture2.Top = 120
                Score = Score + 20
                Picture2.BackColor = vbWhite
                Press(2) = 0
                Hard(2) = False
                Call CheckCounter
            Else
                Let Press(2) = Press(2) + 1
            End If
        Else
            Picture2.Top = 120
            Score = Score + 10
            Call CheckCounter
            Call Block2Hard
        End If
    Else
        sound = sndPlaySound(App.Path + "\buzz.wav", 1)
        Score = Score - 10
        Fdup = True
        If Score < -150 Then
            Call HideSquares
            Call StopG
            lblGO.Visible = True
        End If
    End If
    End If
End If

If KeyCode = vbKeyD Then
    lblD.ForeColor = vbRed
    If Paused = True Or Started = False Then
    Else
    If Picture3.Top >= 2040 And Picture3.Top <= 2570 And Level = 1 Then
        Picture3.Top = 120
        Score = Score + 10
        Call CheckCounter
    ElseIf Picture3.Top >= 2160 And Picture3.Top <= 2570 And Level = 2 Then
        If Hard(3) = True Then
            If Press(3) = 1 Then
                Picture3.Top = 120
                Score = Score + 20
                Picture3.BackColor = vbWhite
                Press(3) = 0
                Hard(3) = False
                Call CheckCounter
            Else
                Let Press(3) = Press(3) + 1
            End If
        Else
            Picture3.Top = 120
            Score = Score + 10
            Call CheckCounter
            Call Block3Hard
        End If
    ElseIf Picture3.Top >= 2280 And Picture3.Top <= 2570 And Level = 3 Then
        If Hard(3) = True Then
            If Press(3) = 1 Then
                Picture3.Top = 120
                Score = Score + 20
                Picture3.BackColor = vbWhite
                Press(3) = 0
                Hard(3) = False
                Call CheckCounter
            Else
                Let Press(3) = Press(3) + 1
            End If
        Else
            Picture3.Top = 120
            Score = Score + 10
            Call CheckCounter
            Call Block3Hard
        End If
    Else
        sound = sndPlaySound(App.Path + "\buzz.wav", 1)
        Score = Score - 10
        Fdup = True
        If Score < -150 Then
            Call HideSquares
            Call StopG
            lblGO.Visible = True
        End If
    End If
    End If
End If

If KeyCode = vbKeyJ Then
    lblJ.ForeColor = vbRed
    If Paused = True Or Started = False Then
    Else
    If Picture4.Top >= 2040 And Picture4.Top <= 2570 And Level = 1 Then
        Picture4.Top = 120
        Score = Score + 10
        Call CheckCounter
    ElseIf Picture4.Top >= 2160 And Picture4.Top <= 2570 And Level = 2 Then
        If Hard(4) = True Then
            If Press(4) = 1 Then
                Picture4.Top = 120
                Score = Score + 20
                Picture4.BackColor = vbWhite
                Press(4) = 0
                Hard(4) = False
                Call CheckCounter
            Else
                Let Press(4) = Press(4) + 1
            End If
        Else
            Picture4.Top = 120
            Score = Score + 10
            Call CheckCounter
            Call Block4Hard
        End If
    ElseIf Picture4.Top >= 2280 And Picture4.Top <= 2570 And Level = 3 Then
        If Hard(4) = True Then
            If Press(4) = 1 Then
                Picture4.Top = 120
                Score = Score + 20
                Picture4.BackColor = vbWhite
                Press(4) = 0
                Hard(4) = False
                Call CheckCounter
            Else
                Let Press(4) = Press(4) + 1
            End If
        Else
            Picture4.Top = 120
            Score = Score + 10
            Call CheckCounter
            Call Block4Hard
        End If
    Else
        sound = sndPlaySound(App.Path + "\buzz.wav", 1)
        Score = Score - 10
        Fdup = True
        If Score < -150 Then
            Call HideSquares
            Call StopG
            lblGO.Visible = True
        End If
    End If
    End If
End If

If KeyCode = vbKeyK Then
    lblK.ForeColor = vbRed
    If Paused = True Or Started = False Then
    Else
    If Picture5.Top >= 2040 And Picture5.Top <= 2570 And Level = 1 Then
        Picture5.Top = 120
        Score = Score + 10
        Call CheckCounter
    ElseIf Picture5.Top >= 2160 And Picture5.Top <= 2570 And Level = 2 Then
        If Hard(5) = True Then
            If Press(5) = 1 Then
                Picture5.Top = 120
                Score = Score + 20
                Picture5.BackColor = vbWhite
                Press(5) = 0
                Hard(5) = False
                Call CheckCounter
            Else
                Let Press(5) = Press(5) + 1
            End If
        Else
            Picture5.Top = 120
            Score = Score + 10
            Call CheckCounter
            Call Block5Hard
        End If
    ElseIf Picture5.Top >= 2280 And Picture5.Top <= 2570 And Level = 3 Then
        If Hard(5) = True Then
            If Press(5) = 1 Then
                Picture5.Top = 120
                Score = Score + 20
                Picture5.BackColor = vbWhite
                Press(5) = 0
                Hard(5) = False
                Call CheckCounter
            Else
                Let Press(5) = Press(5) + 1
            End If
        Else
            Picture5.Top = 120
            Score = Score + 10
            Call CheckCounter
            Call Block5Hard
        End If
    Else
        sound = sndPlaySound(App.Path + "\buzz.wav", 1)
        Score = Score - 10
        Fdup = True
        If Score < -150 Then
            Call HideSquares
            Call StopG
            lblGO.Visible = True
        End If
    End If
    End If
End If

If KeyCode = vbKeyL Then
    lblL.ForeColor = vbRed
    If Paused = True Or Started = False Then
    Else
    If Picture6.Top >= 2040 And Picture6.Top <= 2570 And Level = 1 Then
        Picture6.Top = 120
        Score = Score + 10
        Call CheckCounter
    ElseIf Picture6.Top >= 2160 And Picture6.Top <= 2570 And Level = 2 Then
        If Hard(6) = True Then
            If Press(6) = 1 Then
                Picture6.Top = 120
                Score = Score + 20
                Picture6.BackColor = vbWhite
                Press(6) = 0
                Hard(6) = False
                Call CheckCounter
            Else
                Let Press(6) = Press(6) + 1
            End If
        Else
            Picture6.Top = 120
            Score = Score + 10
            Call CheckCounter
            Call Block6Hard
        End If
    ElseIf Picture6.Top >= 2280 And Picture6.Top <= 2570 And Level = 3 Then
        If Hard(6) = True Then
            If Press(6) = 1 Then
                Picture6.Top = 120
                Score = Score + 20
                Picture6.BackColor = vbWhite
                Press(6) = 0
                Hard(6) = False
                Call CheckCounter
            Else
                Let Press(6) = Press(6) + 1
            End If
        Else
            Picture6.Top = 120
            Score = Score + 10
            Call CheckCounter
            Call Block6Hard
        End If
    Else
        sound = sndPlaySound(App.Path + "\buzz.wav", 1)
        Score = Score - 10
        Fdup = True
        If Score < -150 Then
            Call HideSquares
            Call StopG
            lblGO.Visible = True
        End If
    End If
    End If
End If

If KeyCode = vbKeyP Then
    If Started = True Then
    If Paused = False Then
        HideSquares
        lblPaused.Visible = True
        Paused = True
    Else
        ShowSquares
        lblPaused.Visible = False
        Paused = False
    End If
    End If
End If

If KeyCode = vbKeyReturn And Shift = 3 Then Call Picture7_Click
lblScore.Caption = Score

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyA Then lblA.ForeColor = vbBlack
If KeyCode = vbKeyS Then lblS.ForeColor = vbBlack
If KeyCode = vbKeyD Then lblD.ForeColor = vbBlack
If KeyCode = vbKeyJ Then lblJ.ForeColor = vbBlack
If KeyCode = vbKeyK Then lblK.ForeColor = vbBlack
If KeyCode = vbKeyL Then lblL.ForeColor = vbBlack
End Sub

Private Sub Form_Load()
Speed(1) = 30
Speed(2) = 40
Speed(3) = 25
Speed(4) = 10
Speed(5) = 35
Speed(6) = 20

Started = False
Level = 1
Stage = 1
End Sub

Private Sub mnuArcade_Click()
If Started = False Then
    mnuArcade.Caption = "Stop"
    mnuCustom.Enabled = False
    
    Speed(1) = 30
    Speed(2) = 40
    Speed(3) = 25
    Speed(4) = 10
    Speed(5) = 35
    Speed(6) = 20
    
    Counter = 20
    Line5.Y1 = 2400
    Line5.Y2 = 2400
    Level = 1
    Stage = 1
    lblLevel.Caption = "Level - 1"
    lblStage.Caption = "Stage - 1"
    Call ShowSquares
    Call Start
Else
    mnuArcade.Caption = "Arcade Mode"
    mnuCustom.Enabled = True
    
    Call HideSquares
    Call StopG
End If

End Sub

Private Sub mnuCustom_Click()
If Started = False Then
    Form2.Show
Else
    Custom = False
    mnuCustom.Caption = "Custom Mode"
    mnuArcade.Enabled = True
    Call StopG
    Call HideSquares
End If
End Sub

Private Sub mnuExit_Click()
End
End Sub

Public Sub Picture7_Click()
mnuArcade.Enabled = True

If Level = 1 And Stage = 1 Then
    Picture7.Visible = False
    Stage = 2
    Counter = 25
    lblStage.Caption = "Stage - 2"
    Speed(1) = 40
    Speed(2) = 50
    Speed(3) = 35
    Speed(4) = 20
    Speed(5) = 45
    Speed(6) = 30
    Call Start
    Call ShowSquares
ElseIf Level = 1 And Stage = 2 Then
    Picture7.Visible = False
    Stage = 3
    Counter = 30
    lblStage.Caption = "Stage - 3"
    Speed(1) = 60
    Speed(2) = 45
    Speed(3) = 55
    Speed(4) = 50
    Speed(5) = 40
    Speed(6) = 35
    Call Start
    Call ShowSquares
ElseIf Level = 1 And Stage = 3 Then
    Picture7.Visible = False
    Stage = 1
    Level = 2
    Counter = 35
    lblLevel.Caption = "Level - 2"
    lblStage.Caption = "Stage - 1"
    Speed(1) = 40
    Speed(2) = 30
    Speed(3) = 50
    Speed(4) = 45
    Speed(5) = 55
    Speed(6) = 60
    Line5.Y1 = 2520
    Line5.Y2 = 2520
    Call Start
    Call ShowSquares
ElseIf Level = 2 And Stage = 1 Then
    Picture7.Visible = False
    Stage = 2
    Counter = 40
    lblLevel.Caption = "Level - 2"
    lblStage.Caption = "Stage - 2"
    Speed(1) = 55
    Speed(2) = 65
    Speed(3) = 50
    Speed(4) = 35
    Speed(5) = 60
    Speed(6) = 45
    Call Start
    Call ShowSquares
ElseIf Level = 2 And Stage = 2 Then
    Picture7.Visible = False
    Stage = 3
    Counter = 45
    lblLevel.Caption = "Level - 2"
    lblStage.Caption = "Stage - 3"
    Speed(1) = 60
    Speed(2) = 70
    Speed(3) = 55
    Speed(4) = 40
    Speed(5) = 65
    Speed(6) = 50
    Call Start
    Call ShowSquares
ElseIf Level = 2 And Stage = 3 Then
    Picture7.Visible = False
    Stage = 1
    Level = 3
    Counter = 50
    lblLevel.Caption = "Level - 3"
    lblStage.Caption = "Stage - 1"
    Speed(1) = 65
    Speed(2) = 75
    Speed(3) = 60
    Speed(4) = 45
    Speed(5) = 70
    Speed(6) = 55
    Line5.Y1 = 2650
    Line5.Y2 = 2650
    Call Start
    Call ShowSquares
 ElseIf Level = 3 And Stage = 1 Then
    Picture7.Visible = False
    Stage = 2
    Counter = 55
    lblLevel.Caption = "Level - 3"
    lblStage.Caption = "Stage - 2"
    Speed(1) = 60
    Speed(2) = 70
    Speed(3) = 55
    Speed(4) = 40
    Speed(5) = 65
    Speed(6) = 60
    Call Start
    Call ShowSquares
ElseIf Level = 3 And Stage = 2 Then
    Picture7.Visible = False
    Stage = 3
    Counter = 55
    lblLevel.Caption = "Level - Final"
    lblStage.Caption = "Stage - Final"
    Speed(1) = 65
    Speed(2) = 75
    Speed(3) = 60
    Speed(4) = 45
    Speed(5) = 70
    Speed(6) = 65
    Call Start
    Call ShowSquares
End If
End Sub

Private Sub Timer1_Timer()
Dim sound As Long

Picture1.Top = Picture1.Top + Speed(1)
If Picture1.Top >= 2520 Then
    sound = sndPlaySound(App.Path + "\buzz.wav", 1)
    Score = Score - 10
    Miss = Miss + 1
    Press(1) = 0
    Picture1.Top = 120
    lblScore.Caption = Score
    MissCalc
End If
End Sub

Private Sub Timer2_Timer()
Dim sound As Long

Picture2.Top = Picture2.Top + Speed(2)
If Picture2.Top >= 2520 Then
    sound = sndPlaySound(App.Path + "\buzz.wav", 1)
    Score = Score - 10
    Miss = Miss + 1
    Press(2) = 0
    Picture2.Top = 120
    lblScore.Caption = Score
    MissCalc
End If
End Sub

Private Sub Timer3_Timer()
Dim sound As Long

Picture3.Top = Picture3.Top + Speed(3)
If Picture3.Top >= 2520 Then
    sound = sndPlaySound(App.Path + "\buzz.wav", 1)
    Score = Score - 10
    Miss = Miss + 1
    Press(3) = 0
    Picture3.Top = 120
    lblScore.Caption = Score
    MissCalc
End If
End Sub

Private Sub MissCalc()
If Miss = 1 Then lblMiss1.Visible = True
If Miss = 2 Then lblMiss2.Visible = True
If Miss = 3 Then
    lblMiss3.Visible = True
    Call HideSquares
    Call StopG
    lblGO.Visible = True
End If
End Sub

Private Sub Timer4_Timer()
Dim sound As Long

Picture4.Top = Picture4.Top + Speed(4)
If Picture4.Top >= 2520 Then
    sound = sndPlaySound(App.Path + "\buzz.wav", 1)
    Score = Score - 10
    Miss = Miss + 1
    Press(4) = 0
    Picture4.Top = 120
    lblScore.Caption = Score
    MissCalc
End If

End Sub

Private Sub Timer5_Timer()
Dim sound As Long

Picture5.Top = Picture5.Top + Speed(5)
If Picture5.Top >= 2520 Then
    sound = sndPlaySound(App.Path + "\buzz.wav", 1)
    Score = Score - 10
    Miss = Miss + 1
    Press(5) = 0
    Picture5.Top = 120
    lblScore.Caption = Score
    MissCalc
End If

End Sub

Private Sub Timer6_Timer()
Dim sound As Long

Picture6.Top = Picture6.Top + Speed(6)
If Picture6.Top >= 2520 Then
    sound = sndPlaySound(App.Path + "\buzz.wav", 1)
    Score = Score - 10
    Miss = Miss + 1
    Press(6) = 0
    Picture6.Top = 120
    lblScore.Caption = Score
    MissCalc
End If

End Sub

Private Sub Block1Hard()
Dim ran As Integer
Randomize

ran = Int((5 - 1 + 1) * Rnd + 1)

If ran = 3 Then Hard(1) = True: Picture1.BackColor = vbRed
End Sub

Private Sub Block2Hard()
Dim ran As Integer
Randomize

ran = Int((5 - 1 + 1) * Rnd + 1)

If ran = 3 Then Hard(2) = True: Picture2.BackColor = vbRed
End Sub

Private Sub Block3Hard()
Dim ran As Integer
Randomize

ran = Int((5 - 1 + 1) * Rnd + 1)

If ran = 3 Then Hard(3) = True: Picture3.BackColor = vbRed
End Sub

Private Sub Block4Hard()
Dim ran As Integer
Randomize

ran = Int((5 - 1 + 1) * Rnd + 1)

If ran = 4 Then Hard(4) = True: Picture4.BackColor = vbRed
End Sub

Private Sub Block5Hard()
Dim ran As Integer
Randomize

ran = Int((5 - 1 + 1) * Rnd + 1)

If ran = 5 Then Hard(5) = True: Picture5.BackColor = vbRed
End Sub

Private Sub Block6Hard()
Dim ran As Integer
Randomize

ran = Int((5 - 1 + 1) * Rnd + 1)

If ran = 5 Then Hard(6) = True: Picture6.BackColor = vbRed
End Sub

Private Sub CheckCounter()
Counter = Counter - 1
lblCounter.Caption = Counter
If Counter = 0 Then
    Started = False
    Call HideSquares
    mnuArcade.Enabled = False
    lblMiss1.Visible = False
    lblMiss2.Visible = False
    lblMiss3.Visible = False
    lblLevelO.Visible = True
    If Miss = 0 And Fdup = False Then lblPerfect.Visible = True
    
    If Custom = True Then
        mnuCustom.Caption = "Custom Mode"
        mnuArcade.Enabled = True
    Else
        Picture7.Visible = True
        Picture7.Print " Click here to continue!"
    End If
    
    If Level = 3 And Stage = 3 Then
        lblCongrats.Visible = True
        Picture7.Visible = False
        mnuArcade.Caption = "Arcade Mode"
        mnuArcade.Enabled = True
        mnuCustom.Enabled = True
        Level = 1
        Stage = 1
    End If
    
End If
End Sub

Private Sub Timer7_Timer()
lblA.ForeColor = vbBlack
lblS.ForeColor = vbBlack
lblD.ForeColor = vbBlack
lblJ.ForeColor = vbBlack
lblK.ForeColor = vbBlack
lblL.ForeColor = vbBlack

Picture1.Top = Picture1.Top + Speed(1)
Picture2.Top = Picture2.Top + Speed(2)
Picture3.Top = Picture3.Top + Speed(3)
Picture4.Top = Picture4.Top + Speed(4)
Picture5.Top = Picture5.Top + Speed(5)
Picture6.Top = Picture6.Top + Speed(6)

If Picture1.Top >= 2140 Then Picture1.Top = 120: lblA.ForeColor = vbRed
If Picture2.Top >= 2140 Then Picture2.Top = 120: lblS.ForeColor = vbRed
If Picture3.Top >= 2140 Then Picture3.Top = 120: lblD.ForeColor = vbRed
If Picture4.Top >= 2140 Then Picture4.Top = 120: lblJ.ForeColor = vbRed
If Picture5.Top >= 2140 Then Picture5.Top = 120: lblK.ForeColor = vbRed
If Picture6.Top >= 2140 Then Picture6.Top = 120: lblL.ForeColor = vbRed

End Sub

