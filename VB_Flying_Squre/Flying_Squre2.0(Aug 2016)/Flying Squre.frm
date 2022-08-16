VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "飞翔的方块"
   ClientHeight    =   8565
   ClientLeft      =   4095
   ClientTop       =   2415
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
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   225
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
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   2100
      Index           =   9
      Left            =   12500
      Top             =   6500
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   4500
      Index           =   8
      Left            =   12500
      Top             =   0
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   4100
      Index           =   7
      Left            =   10000
      Top             =   4500
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   2500
      Index           =   6
      Left            =   10000
      Top             =   0
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   4400
      Index           =   5
      Left            =   7500
      Top             =   4200
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   2200
      Index           =   4
      Left            =   7500
      Top             =   0
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   2100
      Index           =   3
      Left            =   5000
      Top             =   6500
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   4500
      Index           =   2
      Left            =   5000
      Top             =   0
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   3600
      Index           =   1
      Left            =   2500
      Top             =   5000
      Width           =   600
   End
   Begin VB.Shape Shape2 
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

Public Sub TimeDelay(ByVal PauseSecond As Single)
 Dim Star, PauseTime
 Star = Timer
 PauseTime = PauseSecond
 Do While Timer < Star + PauseTime
 DoEvents
 Loop
End Sub


Private Sub Form_GotFocus()
t2 = 0
Shape2(0).Left = 2500
Shape2(1).Left = 2500
Shape2(2).Left = 5000
Shape2(3).Left = 5000
Shape2(4).Left = 7500
Shape2(5).Left = 7500
Shape2(6).Left = 10000
Shape2(7).Left = 10000
Shape2(8).Left = 12500
Shape2(9).Left = 12500
Shape2(0).Visible = True
Shape2(1).Visible = True
Shape2(2).Visible = True
Shape2(3).Visible = True
Shape2(4).Visible = True
Shape2(5).Visible = True
Shape2(6).Visible = True
Shape2(7).Visible = True
Shape2(8).Visible = True
Shape2(9).Visible = True
Shape1.Top = 3000: Shape1.Left = 0
If a5 = 0 Then t1 = 40
If a5 = 1 Then t1 = 60
If a5 = 2 Then t1 = 80
If a5 = 3 Then t1 = 110
If a5 = 0 Then Timer3.Interval = 200
If a5 = 1 Then Timer3.Interval = 150
If a5 = 2 Then Timer3.Interval = 75
If a5 = 3 Then Timer3.Interval = 50
If a2 = 0 Then Shape1.FillColor = &H0&
If a2 = 1 Then Shape1.FillColor = &HFF0000
If a2 = 2 Then Shape1.FillColor = &HFF00&
If a2 = 3 Then Shape1.FillColor = &HFF&
If a2 = 4 Then Shape1.FillColor = &HFFFF&
If a3 = 0 Then
   Shape2(0).FillColor = &H0&
   Shape2(1).FillColor = &H0&
   Shape2(2).FillColor = &H0&
   Shape2(3).FillColor = &H0&
   Shape2(4).FillColor = &H0&
   Shape2(5).FillColor = &H0&
   Shape2(6).FillColor = &H0&
   Shape2(7).FillColor = &H0&
   Shape2(8).FillColor = &H0&
   Shape2(9).FillColor = &H0&
End If
If a3 = 1 Then
   Shape2(0).FillColor = &HFF0000
   Shape2(1).FillColor = &HFF0000
   Shape2(2).FillColor = &HFF0000
   Shape2(3).FillColor = &HFF0000
   Shape2(4).FillColor = &HFF0000
   Shape2(5).FillColor = &HFF0000
   Shape2(6).FillColor = &HFF0000
   Shape2(7).FillColor = &HFF0000
   Shape2(8).FillColor = &HFF0000
   Shape2(9).FillColor = &HFF0000
End If
If a3 = 2 Then
   Shape2(0).FillColor = &HFF00&
   Shape2(1).FillColor = &HFF00&
   Shape2(2).FillColor = &HFF00&
   Shape2(3).FillColor = &HFF00&
   Shape2(4).FillColor = &HFF00&
   Shape2(5).FillColor = &HFF00&
   Shape2(6).FillColor = &HFF00&
   Shape2(7).FillColor = &HFF00&
   Shape2(8).FillColor = &HFF00&
   Shape2(9).FillColor = &HFF00&
End If
If a3 = 3 Then
   Shape2(0).FillColor = &HFF&
   Shape2(1).FillColor = &HFF&
   Shape2(2).FillColor = &HFF&
   Shape2(3).FillColor = &HFF&
   Shape2(4).FillColor = &HFF&
   Shape2(5).FillColor = &HFF&
   Shape2(6).FillColor = &HFF&
   Shape2(7).FillColor = &HFF&
   Shape2(8).FillColor = &HFF&
   Shape2(9).FillColor = &HFF&
End If
If a3 = 4 Then
   Shape2(0).FillColor = &HFFFF&
   Shape2(1).FillColor = &HFFFF&
   Shape2(2).FillColor = &HFFFF&
   Shape2(3).FillColor = &HFFFF&
   Shape2(4).FillColor = &HFFFF&
   Shape2(5).FillColor = &HFFFF&
   Shape2(6).FillColor = &HFFFF&
   Shape2(7).FillColor = &HFFFF&
   Shape2(8).FillColor = &HFFFF&
   Shape2(9).FillColor = &HFFFF&
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Timer5.Enabled = True Then
Y = 1
Else
    If KeyCode = 90 And sh(0) > 0 And of(0) = 0 Then '"z"
        PlaySound "E:\VB6.0\VB_Flying_Squre\buffer.wav", 0, 1
        sh(0) = sh(0) - 1
        t1 = t1 - 20
        of(0) = 1
        TimeDelay (5)
        t1 = t1 + 20
        of(0) = 0
    End If
    If KeyCode = 88 And sh(1) > 0 And of(1) = 0 Then '"x"
        PlaySound "E:\VB6.0\VB_Flying_Squre\wall_mechine_start.wav", 0, 1
        sh(1) = sh(1) - 1
        sa = 1
        of(1) = 1
        Shape1.FillColor = &H400040
        TimeDelay (5)
        Shape1.FillColor = &H0
        sa = 0
        of(1) = 0
    End If
    If KeyCode = 38 Then '"up key"
        For Y = 1 To 10
            Sleep 15
            Shape1.Top = Shape1.Top - 50
        Next Y
    End If
    If KeyCode = 40 Then '"down key"
        Shape1.Top = Shape1.Top + 250
    End If
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.Title, "Settings", "sh(0)", sh(0)
SaveSetting App.Title, "Settings", "sh(1)", sh(1)
SaveSetting App.Title, "Settings", "s2", s2
Unload Form2
Unload Form3
Unload Form4
End Sub


Private Sub Timer1_Timer()
If Shape1.Left > 4800 Then
   Shape1.Left = Shape1.Left - t1
   Shape2(0).Left = Shape2(0).Left - t1
   Shape2(1).Left = Shape2(0).Left
   Shape2(2).Left = Shape2(2).Left - t1
   Shape2(3).Left = Shape2(2).Left
   Shape2(4).Left = Shape2(4).Left - t1
   Shape2(5).Left = Shape2(4).Left
   Shape2(6).Left = Shape2(6).Left - t1
   Shape2(7).Left = Shape2(6).Left
   Shape2(8).Left = Shape2(8).Left - t1
   Shape2(9).Left = Shape2(8).Left
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
ha = Rndz(1000, 5000)
If Shape2(0).Left <= -200 Then
   Shape2(0).Left = Shape2(8).Left + 2500
   Shape2(0).Height = ha
   Shape2(1).Height = 7000 - ha
   Shape2(1).Top = Shape2(0).Height + 2000
   Shape2(0).Visible = True
   Shape2(1).Visible = True
ElseIf Shape2(2).Left <= -200 Then
   Shape2(2).Left = Shape2(0).Left + 2500
   Shape2(2).Height = ha
   Shape2(3).Height = 7000 - ha
   Shape2(3).Top = Shape2(2).Height + 2000
   Shape2(2).Visible = True
   Shape2(3).Visible = True
ElseIf Shape2(4).Left <= -200 Then
   Shape2(4).Left = Shape2(2).Left + 2500
   Shape2(4).Height = ha
   Shape2(5).Height = 7000 - ha
   Shape2(5).Top = Shape2(4).Height + 2000
   Shape2(4).Visible = True
   Shape2(5).Visible = True
ElseIf Shape2(6).Left <= -200 Then
   Shape2(6).Left = Shape2(4).Left + 2500
   Shape2(6).Height = ha
   Shape2(7).Height = 7000 - ha
   Shape2(7).Top = Shape2(6).Height + 2000
   Shape2(6).Visible = True
   Shape2(7).Visible = True
ElseIf Shape2(8).Left <= -200 Then
   Shape2(8).Left = Shape2(6).Left + 2500
   Shape2(8).Height = ha
   Shape2(9).Height = 7000 - ha
   Shape2(9).Top = Shape2(8).Height + 2000
   Shape2(8).Visible = True
   Shape2(9).Visible = True
End If
If Shape3.Left <= -200 Then
   Shape3.Visible = False: Shape3.Top = 0 _
   : Shape3.Left = 0: TimeDelay (2): Shape3.Top _
   = Rndz(2000, 3000): Shape3.Left = 6240: Shape3.Visible = True
End If
End Sub
'Edit by Aden
'4/22/2017

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
   Score.Caption = "分数:" + Str(s1)
ElseIf IsTouched(Shape1, Shape3) And Shape3.FillColor = &HC0FFC0 Then
   s1 = s1 + 2: Shape3.Visible = False: Shape3.Top = 0 _
   : Shape3.Left = 0: TimeDelay (4): Shape3.Top = Rndz(2000, 4000): Shape3.Left = 6240: Shape3.Visible = True
   Score.Caption = "分数:" + Str(s1)
ElseIf IsTouched(Shape1, Shape3) And Shape3.FillColor = &HFFFF00 Then
   s1 = s1 + 3: Shape3.Visible = False: Shape3.Top = 0 _
   : Shape3.Left = 0: TimeDelay (3): Shape3.Top = Rndz(2000, 4000): Shape3.Left = 6240: Shape3.Visible = True
   Score.Caption = "分数:" + Str(s1)
ElseIf IsTouched(Shape1, Shape3) And Shape3.FillColor = &HFF Then
   s1 = s1 + 4: Shape3.Visible = False: Shape3.Top = 0 _
   : Shape3.Left = 0: TimeDelay (2): Shape3.Top = Rndz(2000, 4000): Shape3.Left = 6240: Shape3.Visible = True
   Score.Caption = "分数:" + Str(s1)
End If
If IsTouched(Shape1, Shape2(0)) And sa = 0 Or _
IsTouched(Shape1, Shape2(1)) And sa = 0 Or _
IsTouched(Shape1, Shape2(2)) And sa = 0 Or _
IsTouched(Shape1, Shape2(3)) And sa = 0 Or _
IsTouched(Shape1, Shape2(4)) And sa = 0 Or _
IsTouched(Shape1, Shape2(5)) And sa = 0 Or _
IsTouched(Shape1, Shape2(6)) And sa = 0 Or _
IsTouched(Shape1, Shape2(7)) And sa = 0 Or _
IsTouched(Shape1, Shape2(8)) And sa = 0 Or _
IsTouched(Shape1, Shape2(9)) And sa = 0 Then
   PlaySound "E:\VB6.0\VB_Flying_Squre\bomb.wav", 0, 1 _
   : Sleep 1000: Timer1.Enabled = False: Timer2.Enabled = False
   Timer3.Enabled = False
   Form1.Hide: Form3.Show
End If
If IsTouched(Shape1, Shape2(0)) And sa = 1 Then
    Shape2(0).Visible = False: PlaySound "E:\VB6.0\VB_Flying_Squre\wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Shape2(1)) And sa = 1 Then
    Shape2(1).Visible = False: PlaySound "E:\VB6.0\VB_Flying_Squre\wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Shape2(2)) And sa = 1 Then
    Shape2(2).Visible = False: PlaySound "E:\VB6.0\VB_Flying_Squre\wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Shape2(3)) And sa = 1 Then
    Shape2(3).Visible = False: PlaySound "E:\VB6.0\VB_Flying_Squre\wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Shape2(4)) And sa = 1 Then
    Shape2(4).Visible = False: PlaySound "E:\VB6.0\VB_Flying_Squre\wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Shape2(5)) And sa = 1 Then
    Shape2(5).Visible = False: PlaySound "E:\VB6.0\VB_Flying_Squre\wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Shape2(6)) And sa = 1 Then
    Shape2(6).Visible = False: PlaySound "E:\VB6.0\VB_Flying_Squre\wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Shape2(7)) And sa = 1 Then
    Shape2(7).Visible = False: PlaySound "E:\VB6.0\VB_Flying_Squre\wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Shape2(8)) And sa = 1 Then
    Shape2(8).Visible = False: PlaySound "E:\VB6.0\VB_Flying_Squre\wall_broken.wav", 0, 1
ElseIf IsTouched(Shape1, Shape2(9)) And sa = 1 Then
    Shape2(9).Visible = False: PlaySound "E:\VB6.0\VB_Flying_Squre\wall_broken.wav", 0, 1
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
If a4 = 1 Then Label1 = 3
If a4 = 2 Then Label1 = 2
If a4 = 3 Then Label1 = 1
If a4 = 4 And la = 0 Then Label1 = "开始"
If a4 = 4 And la = 1 Then Label1 = "开始"
If a4 = 5 Then Timer1.Enabled = True: Timer2.Enabled = True: _
Timer3.Enabled = True: Timer5.Enabled = False: Label1.Visible = False: Shape3.Visible = True
a4 = a4 + 1
End Sub
'Edited by Aden
'2016/8/21


