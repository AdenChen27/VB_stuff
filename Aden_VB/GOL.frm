VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8565
   ClientLeft      =   2340
   ClientTop       =   810
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9
   ScaleMode       =   0  'User
   ScaleWidth      =   15
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   6870
      Top             =   4065
   End
   Begin VB.Frame Frame1 
      Height          =   8535
      Left            =   -30
      TabIndex        =   1
      Top             =   45
      Width           =   14865
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   4905
         TabIndex        =   2
         Top             =   1665
         Width           =   4095
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Red:4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   330
      TabIndex        =   4
      Top             =   8025
      Width           =   1995
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      Height          =   6345
      Left            =   1440
      TabIndex        =   3
      Top             =   645
      Width           =   11325
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Red:4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12660
      TabIndex        =   0
      Top             =   8025
      Width           =   1995
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   0
      X2              =   7
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   476
      Left            =   10934
      Top             =   2855
      Width           =   497
   End
   Begin VB.Line Line2line 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   11
      X2              =   11.5
      Y1              =   3.5
      Y2              =   3.5
   End
   Begin VB.Line line1line 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0.5
      X2              =   1
      Y1              =   2.5
      Y2              =   2.5
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   3
      X2              =   10
      Y1              =   1
      Y2              =   1
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   476
      Left            =   497
      Top             =   1903
      Width           =   497
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pl1 As Integer
Public pl2 As Integer
Public gh As Integer
Public rh As Integer
Public play1 As String
Public play2 As String

Private Sub Command1_Click()
Frame1.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 38
    Shape1.Top = Shape1.Top - 0.5
    line1line.X1 = Shape1.Left
    line1line.X2 = Shape1.Left + 0.5
    line1line.Y1 = Shape1.Top
    line1line.Y2 = Shape1.Top
    pl1 = 0
Case 87
    Shape2.Top = Shape2.Top - 0.5
    Line2line.X1 = Shape2.Left
    Line2line.X2 = Shape2.Left + 0.5
    Line2line.Y1 = Shape2.Top
    Line2line.Y2 = Shape2.Top
    pl2 = 0
Case 39
    Shape1.Left = Shape1.Left + 0.5
    line1line.Y1 = Shape1.Top
    line1line.Y2 = Shape1.Top + 0.5
    line1line.X1 = Shape1.Left + 0.5
    line1line.X2 = Shape1.Left + 0.5
    pl1 = 3
Case 68
    Shape2.Left = Shape2.Left + 0.5
    Line2line.Y1 = Shape2.Top
    Line2line.Y2 = Shape2.Top + 0.5
    Line2line.X1 = Shape2.Left + 0.5
    Line2line.X2 = Shape2.Left + 0.5
    pl2 = 3
Case 40
    Shape1.Top = Shape1.Top + 0.5
    line1line.X1 = Shape1.Left
    line1line.X2 = Shape1.Left + 0.5
    line1line.Y1 = Shape1.Top + 0.5
    line1line.Y2 = Shape1.Top + 0.5
    pl1 = 1
Case 83
    Shape2.Top = Shape2.Top + 0.5
    Line2line.X1 = Shape2.Left
    Line2line.X2 = Shape2.Left + 0.5
    Line2line.Y1 = Shape2.Top + 0.5
    Line2line.Y2 = Shape2.Top + 0.5
    pl2 = 1
Case 37
    Shape1.Left = Shape1.Left - 0.5
    line1line.Y1 = Shape1.Top
    line1line.Y2 = Shape1.Top + 0.5
    line1line.X1 = Shape1.Left
    line1line.X2 = Shape1.Left
    pl1 = 2
Case 65
    Shape2.Left = Shape2.Left - 0.5
    Line2line.Y1 = Shape2.Top
    Line2line.Y2 = Shape2.Top + 0.5
    Line2line.X1 = Shape2.Left
    Line2line.X2 = Shape2.Left
    pl2 = 2
End Select
If KeyCode = 191 And Shape1.FillColor <> &HC000C0 Then
    If pl1 = 0 Then
        Line1.Y1 = 0
        Line1.Y2 = Shape1.Top
        Line1.X1 = Shape1.Left + 0.25
        Line1.X2 = Line1.X1
        Line1.Visible = True
        If Line1.X1 > Shape2.Left And Line1.X1 < Shape2.Left + 0.5 Then rh = rh - 1: Label2.Caption = play2 & " " & rh
        If rh = 0 Then Label3.Caption = "The Winner Is " & play1
        TimeDelay (0.5)
        If rh <= 0 Then Unload Form1: Form1.Show
        Line1.Visible = False
    ElseIf pl1 = 1 Then
        Line1.Y1 = Shape1.Top + 0.5
        Line1.Y2 = 9
        Line1.X1 = Shape1.Left + 0.25
        Line1.X2 = Line1.X1
        Line1.Visible = True
        If Line1.X1 > Shape2.Left And Line1.X1 < Shape2.Left + 0.5 Then rh = rh - 1: Label2.Caption = play2 & " " & rh
        If rh = 0 Then Label3.Caption = "The Winner Is " & play1
        TimeDelay (0.5)
        If rh <= 0 Then Unload Form1: Form1.Show
        Line1.Visible = False
    ElseIf pl1 = 2 Then
        Line1.Y1 = Shape1.Top + 0.25
        Line1.Y2 = Shape1.Top + 0.25
        Line1.X1 = 0
        Line1.X2 = Shape1.Left
        Line1.Visible = True
        If Line1.Y1 > Shape2.Top And Line1.Y1 < Shape2.Top + 0.5 Then rh = rh - 1: Label2.Caption = play2 & " " & rh
        If rh = 0 Then Label3.Caption = "The Winner Is " & play1
        TimeDelay (0.5)
        If rh <= 0 Then Unload Form1: Form1.Show
        Line1.Visible = False
    ElseIf pl1 = 3 Then
        Line1.Y1 = Shape1.Top + 0.25
        Line1.Y2 = Shape1.Top + 0.25
        Line1.X1 = Shape1.Left + 0.5
        Line1.X2 = 15
        Line1.Visible = True
        If Line1.Y1 > Shape2.Top And Line1.Y1 < Shape2.Top + 0.5 Then rh = rh - 1: Label2.Caption = play2 & " " & rh
        If rh = 0 Then Label3.Caption = "The Winner Is " & play1
        TimeDelay (0.5)
        If rh <= 0 Then Unload Form1: Form1.Show
        Line1.Visible = False
    End If
End If
If KeyCode = 49 And Shape2.FillColor <> &HC000C0 Then
    If pl2 = 0 Then
        Line2.Y1 = 0
        Line2.Y2 = Shape2.Top
        Line2.X1 = Shape2.Left + 0.25
        Line2.X2 = Line2.X1
        Line2.Visible = True
        If Line2.X1 > Shape1.Left And Line2.X1 < Shape1.Left + 0.5 Then gh = gh - 1: Label1.Caption = play1 & " " & gh
        If gh = 0 Then Label3.Caption = "The Winner Is " & play2
        TimeDelay (0.5)
        If gh <= 0 Then Unload Form1: Form1.Show
        Line2.Visible = False
    ElseIf pl2 = 1 Then
        Line2.Y1 = Shape2.Top + 0.5
        Line2.Y2 = 9
        Line2.X1 = Shape2.Left + 0.25
        Line2.X2 = Line2.X1
        Line2.Visible = True
        If Line2.X1 > Shape1.Left And Line2.X1 < Shape1.Left + 0.5 Then gh = gh - 1: Label1.Caption = play1 & " " & gh
        If gh = 0 Then Label3.Caption = "The Winner Is " & play2
        TimeDelay (0.5)
        If gh <= 0 Then Unload Form1: Form1.Show
        Line2.Visible = False
    ElseIf pl2 = 2 Then
        Line2.Y1 = Shape2.Top + 0.25
        Line2.Y2 = Shape2.Top + 0.25
        Line2.X1 = 0
        Line2.X2 = Shape2.Left
        Line2.Visible = True
        If Line2.Y1 > Shape1.Top And Line2.Y1 < Shape1.Top + 0.5 Then gh = gh - 1: Label1.Caption = play1 & " " & gh
        If gh = 0 Then Label3.Caption = "The Winner Is " & play2
        TimeDelay (0.5)
        If gh <= 0 Then Unload Form1: Form1.Show
        Line2.Visible = False
    ElseIf pl2 = 3 Then
        Line2.Y1 = Shape2.Top + 0.25
        Line2.Y2 = Shape2.Top + 0.25
        Line2.X1 = Shape2.Left + 0.5
        Line2.X2 = 15
        Line2.Visible = True
        If Line2.Y1 > Shape1.Top And Line2.Y1 < Shape1.Top + 0.5 Then gh = gh - 1: Label1.Caption = play1 & " " & gh
        If gh = 0 Then Label3.Caption = "The Winner Is " & play2
        TimeDelay (0.5)
        If gh <= 0 Then Unload Form1: Form1.Show
        Line2.Visible = False
    End If
End If
If KeyCode = 190 Then
    Shape2.FillColor = &HC000C0
    Timer2.Enabled = True
ElseIf KeyCode = 192 Then
    Shape1.FillColor = &HC000C0
    Timer2.Enabled = True
End If
End Sub

Private Sub Form_Load()
pl1 = 1
pl2 = 2
gh = 5
rh = 5
play1 = "P1"
play2 = "P2"
Label1.Caption = play1 & " " & gh
Label2.Caption = play2 & " " & rh
End Sub

Public Sub TimeDelay(ByVal PauseSecond As Single)
 Dim Star, PauseTime
 Star = Timer
 PauseTime = PauseSecond
 Do While Timer < Star + PauseTime
 DoEvents
 Loop
End Sub

Private Sub Timer1_Timer()
If Shape1.Top < 0 Then Shape1.Top = 0
If Shape2.Top < 0 Then Shape2.Top = 0
If Shape1.Top > 8.5 Then Shape1.Top = 8.5
If Shape2.Top > 8.5 Then Shape2.Top = 8.5
If Shape1.Left < 0 Then Shape1.Left = 0
If Shape2.Left < 0 Then Shape2.Left = 0
If Shape1.Left > 14.5 Then Shape1.Left = 14.5
If Shape2.Left > 14.5 Then Shape2.Left = 14.5
End Sub

Private Sub Timer2_Timer()
Shape1.FillColor = &HFF00&
Shape2.FillColor = &HFF&
Timer2.Enabled = False
End Sub
