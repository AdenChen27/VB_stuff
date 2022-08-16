VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FIR"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9330
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text4 
      Height          =   435
      Left            =   7245
      TabIndex        =   4
      Top             =   1470
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox Text3 
      Height          =   435
      Left            =   7245
      TabIndex        =   3
      Top             =   2100
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   9030
      Top             =   0
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   7245
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   7245
      TabIndex        =   1
      Top             =   210
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      FillColor       =   &H8000000F&
      ForeColor       =   &H8000000F&
      Height          =   7000
      Left            =   0
      ScaleHeight     =   11.683
      ScaleMode       =   0  'User
      ScaleWidth      =   11.683
      TabIndex        =   0
      Top             =   0
      Width           =   7000
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         FillStyle       =   0  'Solid
         Height          =   597
         Index           =   0
         Left            =   210
         Shape           =   3  'Circle
         Top             =   2520
         Visible         =   0   'False
         Width           =   597
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Next :"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   7140
      TabIndex        =   5
      Top             =   630
      Width           =   2160
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   5565
      Left            =   7245
      Top             =   1470
      Width           =   1875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
P1.Scale (0, 0)-(12, 12)
End Sub

Private Sub Label1_Click()
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Shape2.Height = 3500
Shape2.Top = 3500
Label1.Top = 2600
End Sub

Private Sub p1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1 = X
Text2 = 12 - Y
Text4 = Fix(X + 0.5)
Text3 = Fix(12 - Y + 0.5)

If Fix(X + 0.5) = 0 Or Fix(X + 0.5) = 12 Or Fix(12 - Y + 0.5) = 0 Or Fix(12 - Y + 0.5) = 12 Then
Else
    If ti = 0 Then
       Shape1(0).Top = 11 - (Fix(12 - Y + 0.5) - 0.5)
       Shape1(0).Left = (Fix(X + 0.5) - 0.5)
       Shape1(0).Visible = True
       ti = 1
       b(Fix(X + 0.5), Fix(12 - Y + 0.5), co) = 1
       co = 1
       Shape2.FillColor = &HFFFFFF
    ElseIf ti >= 1 And b(Text4.Text, Text3.Text, 0) = 0 And b(Text4.Text, Text3.Text, 1) = 0 Then
       Load Shape1(ti)
       If co = 0 Then
          Shape1(ti).FillColor = &H0
       ElseIf co = 1 Then
          Shape1(ti).FillColor = &HFFFFFF
       End If
       Shape1(ti).Top = 11 - (Fix(12 - Y + 0.5) - 0.5)
       Shape1(ti).Left = (Fix(X + 0.5) - 0.5)
       Shape1(ti).Visible = True
       ti = ti + 1
       b(Fix(X + 0.5), Fix(12 - Y + 0.5), co) = 1
       If co = 0 Then
          co = 1
          Shape2.FillColor = &HFFFFFF
       ElseIf co = 1 Then
          co = 0
          Shape2.FillColor = &H0
       End If
    End If
End If
End Sub

Private Sub Timer1_Timer()
If a(0) = 0 Then
    For X = 1 To 11
       a(1) = a(1) + 1
       P1.Line (1, a(1))-(11, a(1)), &H0
    Next X
    a(1) = 0
    For X = 1 To 11
       a(1) = a(1) + 1
       P1.Line (a(1), 1)-(a(1), 11), &H0
    Next X
    Timer1.Enabled = False
    a(0) = 1
ElseIf a(0) = 1 Then End
End If
End Sub
