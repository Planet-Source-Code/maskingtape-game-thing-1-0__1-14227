Attribute VB_Name = "Module1"
Option Explicit

Global Fdup As Boolean
Global Started As Boolean
Global Level As Integer
Global Stage As Integer
Global Paused As Boolean
Global Counter As Integer
Global Miss As Integer
Global Score As Integer
Global Custom As Boolean

Global Hard(1 To 6)
Global Press(1 To 6)
Global Speed(1 To 6)

Public Sub HideSquares()
Form1.Timer1.Enabled = False
Form1.Timer2.Enabled = False
Form1.Timer3.Enabled = False
Form1.Timer4.Enabled = False
Form1.Timer5.Enabled = False
Form1.Timer6.Enabled = False

Form1.Picture1.Visible = False
Form1.Picture2.Visible = False
Form1.Picture3.Visible = False
Form1.Picture4.Visible = False
Form1.Picture5.Visible = False
Form1.Picture6.Visible = False
End Sub

Public Sub ShowSquares()
Form1.Timer1.Enabled = True
Form1.Timer2.Enabled = True
Form1.Timer3.Enabled = True
Form1.Timer4.Enabled = True
Form1.Timer5.Enabled = True
Form1.Timer6.Enabled = True

Form1.Picture1.Visible = True
Form1.Picture2.Visible = True
Form1.Picture3.Visible = True
Form1.Picture4.Visible = True
Form1.Picture5.Visible = True
Form1.Picture6.Visible = True

End Sub

Public Sub Start()
Started = True
Score = 0
Miss = 0

Fdup = False
Hard(1) = False
Hard(2) = False
Hard(3) = False
Hard(4) = False
Hard(5) = False
Hard(6) = False
Press(1) = 0
Press(2) = 0
Press(3) = 0
Press(4) = 0
Press(5) = 0
Press(6) = 0

Form1.Picture1.BackColor = vbWhite
Form1.Picture2.BackColor = vbWhite
Form1.Picture3.BackColor = vbWhite
Form1.Picture4.BackColor = vbWhite
Form1.Picture5.BackColor = vbWhite
Form1.Picture6.BackColor = vbWhite

Form1.lblScore.Caption = Score
Form1.lblCounter.Caption = Counter
Form1.Picture1.Top = 120
Form1.Picture2.Top = 120
Form1.Picture3.Top = 120
Form1.Picture4.Top = 120
Form1.Picture5.Top = 120
Form1.Picture6.Top = 120
Form1.lblGO.Visible = False
Form1.lblLevelO.Visible = False
Form1.lblPerfect.Visible = False
Form1.lblGameThing.Visible = False
Form1.lblCongrats.Visible = False
Form1.Timer7.Enabled = False

End Sub

Public Sub StopG()
Form1.mnuArcade.Caption = "Arcade Mode"
Form1.mnuArcade.Enabled = True
Form1.mnuCustom.Caption = "Custom Mode"
Form1.mnuCustom.Enabled = True

Started = False
Score = 0
Counter = 0
Miss = 0
Level = 1

Paused = False
Fdup = False
Hard(1) = False
Hard(2) = False
Hard(3) = False
Hard(4) = False
Hard(5) = False
Hard(6) = False
Press(1) = 0
Press(2) = 0
Press(3) = 0
Press(4) = 0
Press(5) = 0
Press(6) = 0

Form1.Picture1.BackColor = vbWhite
Form1.Picture2.BackColor = vbWhite
Form1.Picture3.BackColor = vbWhite
Form1.Picture4.BackColor = vbWhite
Form1.Picture5.BackColor = vbWhite
Form1.Picture6.BackColor = vbWhite
Form1.lblScore.Caption = Score
Form1.lblCounter.Caption = Counter
Form1.Line5.Y1 = 2400
Form1.Line5.Y2 = 2400
Form1.lblMiss1.Visible = False
Form1.lblMiss2.Visible = False
Form1.lblMiss3.Visible = False
Form1.lblLevelO.Visible = False
Form1.lblPerfect.Visible = False
Form1.lblCongrats.Visible = False
Form1.lblPaused.Visible = False
End Sub

