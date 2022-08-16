VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "飞翔的方块"
   ClientHeight    =   8565
   ClientLeft      =   1995
   ClientTop       =   855
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Flying Squre.frx":0000
   ScaleHeight     =   8565
   ScaleWidth      =   14910
   Begin VB.Timer Timer5 
      Interval        =   500
      Left            =   0
      Top             =   1560
   End
   Begin VB.Timer Timer3 
      Interval        =   300
      Left            =   0
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   1500
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   120
   End
   Begin VB.Label Score 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   35.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   705
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   360
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   10800
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   99.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1995
      Left            =   6480
      TabIndex        =   0
      Top             =   2880
      Width           =   1020
   End
   Begin VB.Shape Obstacle 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   2100
      Index           =   9
      Left            =   12500
      Top             =   6500
      Width           =   600
   End
   Begin VB.Shape Obstacle 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   4500
      Index           =   8
      Left            =   12500
      Top             =   0
      Width           =   600
   End
   Begin VB.Shape Obstacle 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   4100
      Index           =   7
      Left            =   10000
      Top             =   4500
      Width           =   600
   End
   Begin VB.Shape Obstacle 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   2500
      Index           =   6
      Left            =   10000
      Top             =   0
      Width           =   600
   End
   Begin VB.Shape Obstacle 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   4400
      Index           =   5
      Left            =   7500
      Top             =   4200
      Width           =   600
   End
   Begin VB.Shape Obstacle 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   2200
      Index           =   4
      Left            =   7500
      Top             =   0
      Width           =   600
   End
   Begin VB.Shape Obstacle 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   2100
      Index           =   3
      Left            =   5000
      Top             =   6500
      Width           =   600
   End
   Begin VB.Shape Obstacle 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   4500
      Index           =   2
      Left            =   5000
      Top             =   0
      Width           =   600
   End
   Begin VB.Shape Obstacle 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   3600
      Index           =   1
      Left            =   2500
      Top             =   5000
      Width           =   600
   End
   Begin VB.Shape Obstacle 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   0
      Left            =   2500
      Top             =   0
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   1000
      Left            =   0
      Top             =   2040
      Width           =   1000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long 'wav sound
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim WithEvents txtTotal As TextBox
Attribute txtTotal.VB_VarHelpID = -1
Dim sc As Integer

Private Sub Form_GotFocus()
Obstacle(0).Left = 2500
Obstacle(1).Left = 2500
Obstacle(2).Left = 5000
Obstacle(3).Left = 5000
Obstacle(4).Left = 7500
Obstacle(5).Left = 7500
Obstacle(6).Left = 10000
Obstacle(7).Left = 10000
Obstacle(8).Left = 12500
Obstacle(9).Left = 12500
For i = 0 To 9
    Obstacle(i).Visible = True
Next i
Shape1.Top = 3000: Shape1.Left = 0
If Grade = 0 Then t1 = 40: Timer3.Interval = 200
If Grade = 1 Then t1 = 60: Timer3.Interval = 150
If Grade = 2 Then t1 = 80: Timer3.Interval = 75
If Grade = 3 Then t1 = 110: Timer3.Interval = 50
If Color_Of_Squre < 0 Then Color_Of_Squre = vbBlack
Shape1.FillColor = Color_Of_Squre
If Color_Of_Obstacle < 0 Then Color_Of_Obstacle = vbBlack
For i = 0 To 9:
    Obstacle(i).FillColor = Color_Of_Obstacle
    Obstacle(i).Visible = True
Next i
If Lang = 0 Then Score.Caption = "分数:"
End Sub
'Edited By Aden An Chen
'Apr 22,2017

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Timer5.Enabled = True Then
    Y = 1
Else
    If KeyCode = 90 And sh(0) > 0 And of(0) = 0 Then '"z"
        PlaySound File_of_Sound & "buffer.wav", 0, 1
        sh(0) = sh(0) - 1
        t1 = t1 - 20
        of(0) = 1
        TimeDelay (5)
        t1 = t1 + 20
        of(0) = 0
    End If
    If KeyCode = 88 And sh(1) > 0 And of(1) = 0 Then '"x"
        PlaySound File_of_Sound & "wall_mechine_start.wav", 0, 1
        sh(1) = sh(1) - 1
        sa = 1
        of(1) = 1
        Shape1.FillColor = &H400040
        TimeDelay (5)
        Shape1.FillColor = Color_Of_Squre
        sa = 0
        of(1) = 0
    End If
    If KeyCode = 38 Then '"up key"
        For Y = 1 To 10
            Sleep 15
            Shape1.Top = Shape1.Top - 75
        Next Y
    End If
    If KeyCode = 40 Then '"down key"
        Shape1.Top = Shape1.Top + 250
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Unload_All
End Sub


Private Sub Timer1_Timer()
If Shape1.Left > 4800 Then
   Shape1.Left = Shape1.Left - t1
   Obstacle(0).Left = Obstacle(0).Left - t1
   Obstacle(1).Left = Obstacle(0).Left
   Obstacle(2).Left = Obstacle(2).Left - t1
   Obstacle(3).Left = Obstacle(2).Left
   Obstacle(4).Left = Obstacle(4).Left - t1
   Obstacle(5).Left = Obstacle(4).Left
   Obstacle(6).Left = Obstacle(6).Left - t1
   Obstacle(7).Left = Obstacle(6).Left
   Obstacle(8).Left = Obstacle(8).Left - t1
   Obstacle(9).Left = Obstacle(8).Left
   Shape3.Left = Shape3.Left - t1
End If
Shape1.Top = Shape1.Top + t1
Shape1.Left = Shape1.Left + t1 - t1 / 4
End Sub

Function IsTouched(a, b) As Boolean
IsTouched = b.Visible And _
Not (a.Left > b.Left + b.Width Or _
     b.Left > a.Left + a.Width Or _
     a.Top > b.Top + b.Height Or _
     b.Top > a.Top + a.Height)
End Function

Private Sub Timer2_Timer()
h = Rndz(1000, 5000) 'hight of the first obstacle
If Obstacle(0).Left <= -200 Then
   Obstacle(0).Left = Obstacle(8).Left + 2500
   Obstacle(0).Height = h
   Obstacle(1).Height = 7000 - h
   Obstacle(1).Top = Obstacle(0).Height + 2000
   Obstacle(0).Visible = True
   Obstacle(1).Visible = True
ElseIf Obstacle(2).Left <= -200 Then
   Obstacle(2).Left = Obstacle(0).Left + 2500
   Obstacle(2).Height = h
   Obstacle(3).Height = 7000 - h
   Obstacle(3).Top = Obstacle(2).Height + 2000
   Obstacle(2).Visible = True
   Obstacle(3).Visible = True
ElseIf Obstacle(4).Left <= -200 Then
   Obstacle(4).Left = Obstacle(2).Left + 2500
   Obstacle(4).Height = h
   Obstacle(5).Height = 7000 - h
   Obstacle(5).Top = Obstacle(4).Height + 2000
   Obstacle(4).Visible = True
   Obstacle(5).Visible = True
ElseIf Obstacle(6).Left <= -200 Then
   Obstacle(6).Left = Obstacle(4).Left + 2500
   Obstacle(6).Height = h
   Obstacle(7).Height = 7000 - h
   Obstacle(7).Top = Obstacle(6).Height + 2000
   Obstacle(6).Visible = True
   Obstacle(7).Visible = True
ElseIf Obstacle(8).Left <= -200 Then
   Obstacle(8).Left = Obstacle(6).Left + 2500
   Obstacle(8).Height = h
   Obstacle(9).Height = 7000 - h
   Obstacle(9).Top = Obstacle(8).Height + 2000
   Obstacle(8).Visible = True
   Obstacle(9).Visible = True
End If
If Shape3.Left <= -200 Then
   Shape3.Visible = False: Shape3.Top = 0 _
   : Shape3.Left = 0: TimeDelay (2): Shape3.Top _
   = Rndz(2000, 3000): Shape3.Left = 6240: Shape3.Visible = True
End If
End Sub
'Edited By Aden An Chen
'Apr 22, 2017

Private Function Rndz(a As Long, b As Long)
    Randomize
    Rndz = Int((a - b + 1) * Rnd() + b)
End Function

Private Sub Timer3_Timer()
If s1 >= 10 Then
   Shape3.FillColor = &HC0FFC0
ElseIf s1 >= 20 Then
   Shape3.FillColor = &HFFFF00
ElseIf s1 >= 30 Then
   Shape3.FillColor = &HFF
End If
If IsTouched(Shape1, Shape3) And Shape3.FillColor = &HC0FFFF Then
   s1 = s1 + 1: Shape3.Visible = False: Shape3.Top = 0 _
   : Shape3.Left = 0: TimeDelay (5): Shape3.Top = Rndz(2000, 4000): Shape3.Left = 6240: Shape3.Visible = True
   If Lang = 0 Then Score.Caption = "分数:" + Str(s1) Else: Score.Caption = "Score:" + Str(s1)
ElseIf IsTouched(Shape1, Shape3) And Shape3.FillColor = &HC0FFC0 Then
   s1 = s1 + 2: Shape3.Visible = False: Shape3.Top = 0 _
   : Shape3.Left = 0: TimeDelay (4): Shape3.Top = Rndz(2000, 4000): Shape3.Left = 6240: Shape3.Visible = True
   If Lang = 0 Then Score.Caption = "分数:" + Str(s1) Else: Score.Caption = "Score:" + Str(s1)
ElseIf IsTouched(Shape1, Shape3) And Shape3.FillColor = &HFFFF00 Then
   s1 = s1 + 3: Shape3.Visible = False: Shape3.Top = 0 _
   : Shape3.Left = 0: TimeDelay (3): Shape3.Top = Rndz(2000, 4000): Shape3.Left = 6240: Shape3.Visible = True
   If Lang = 0 Then Score.Caption = "分数:" + Str(s1) Else: Score.Caption = "Score:" + Str(s1)
ElseIf IsTouched(Shape1, Shape3) And Shape3.FillColor = &HFF Then
   s1 = s1 + 4: Shape3.Visible = False: Shape3.Top = 0 _
   : Shape3.Left = 0: TimeDelay (2): Shape3.Top = Rndz(2000, 4000): Shape3.Left = 6240: Shape3.Visible = True
  If Lang = 0 Then Score.Caption = "分数:" + Str(s1) Else: Score.Caption = "Score:" + Str(s1)
End If
If IsTouched(Shape1, Obstacle(0)) And sa = 0 Or _
IsTouched(Shape1, Obstacle(1)) And sa = 0 Or _
IsTouched(Shape1, Obstacle(2)) And sa = 0 Or _
IsTouched(Shape1, Obstacle(3)) And sa = 0 Or _
IsTouched(Shape1, Obstacle(4)) And sa = 0 Or _
IsTouched(Shape1, Obstacle(5)) And sa = 0 Or _
IsTouched(Shape1, Obstacle(6)) And sa = 0 Or _
IsTouched(Shape1, Obstacle(7)) And sa = 0 Or _
IsTouched(Shape1, Obstacle(8)) And sa = 0 Or _
IsTouched(Shape1, Obstacle(9)) And sa = 0 Then
   PlaySound File_of_Sound & "bomb.wav", 0, 1 _
   : Sleep 1000: Timer1.Enabled = False: Timer2.Enabled = False
   Timer3.Enabled = False
   Form1.Hide: Form3.Show
End If
If IsTouched(Shape1, Obstacle(0)) And sa = 1 Then
    Obstacle(0).Visible = False: PlaySound File_of_Sound & "wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Obstacle(1)) And sa = 1 Then
    Obstacle(1).Visible = False: PlaySound File_of_Sound & "wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Obstacle(2)) And sa = 1 Then
    Obstacle(2).Visible = False: PlaySound File_of_Sound & "wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Obstacle(3)) And sa = 1 Then
    Obstacle(3).Visible = False: PlaySound File_of_Sound & "wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Obstacle(4)) And sa = 1 Then
    Obstacle(4).Visible = False: PlaySound File_of_Sound & "wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Obstacle(5)) And sa = 1 Then
    Obstacle(5).Visible = False: PlaySound File_of_Sound & "wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Obstacle(6)) And sa = 1 Then
    Obstacle(6).Visible = False: PlaySound File_of_Sound & "wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Obstacle(7)) And sa = 1 Then
    Obstacle(7).Visible = False: PlaySound File_of_Sound & "wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Obstacle(8)) And sa = 1 Then
    Obstacle(8).Visible = False: PlaySound File_of_Sound & "wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Obstacle(9)) And sa = 1 Then
    Obstacle(9).Visible = False: PlaySound File_of_Sound & "wall_broken.wav", 0, 1
End If
If Shape1.Top <= 0 Then
   Shape1.Top = Shape1.Top + 300
End If
If Shape1.Top > 9000 Then
    Form1.Hide: Form3.Show
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
End If
End Sub

Private Sub Timer5_Timer()
If Count_Down = 1 Then Label1 = 3
If Count_Down = 2 Then Label1 = 2
If Count_Down = 3 Then Label1 = 1
If Count_Down = 5 Then
    If Lang = 0 Then Label1 = "开始" Else: Label1 = "Start"
End If
If Count_Down = 5 Then TimeDelay (1): Timer1.Enabled = True: Timer2.Enabled = True: _
Timer3.Enabled = True: Timer5.Enabled = False: Label1.Visible = False: Shape3.Visible = True
Count_Down = Count_Down + 1
End Sub

