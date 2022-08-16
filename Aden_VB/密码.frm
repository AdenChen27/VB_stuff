VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4740
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1050
      TabIndex        =   2
      Top             =   1665
      Width           =   2340
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   555
      Left            =   60
      TabIndex        =   1
      Text            =   "Password"
      Top             =   825
      Width           =   4380
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Text            =   "Username"
      Top             =   195
      Width           =   4380
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   315
      Top             =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username or Password was wrong."
      Height          =   240
      Left            =   840
      TabIndex        =   3
      Top             =   1380
      Visible         =   0   'False
      Width           =   3435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Pass As Integer

Private Sub Command1_Click()
Label1.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And X < 300 And Y < 300 And Pass = 1 Then Pass = 2
End Sub

Private Sub Form_Load()
For X = 0 To Right(Time, 2)
    a = Rnd
Next X
Timer1.Interval = Rnd * 5000
If Timer1.Interval < 2000 Then Timer1.Interval = Timer1.Interval + 2000
Form1.BackColor = RGB(236, 236, 236)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Pass = 2 Then Form2.Show
End Sub

Private Sub Text1_Click()
Text1.Text = ""
Text1.ForeColor = vbBlack
Text1.PasswordChar = "*"
Label1.Visible = False
End Sub

Private Sub Text2_Click()
Text2.Text = ""
Text2.ForeColor = vbBlack
Text2.PasswordChar = "*"
Label1.Visible = False
End Sub

Private Sub Timer1_Timer()
If Timer1.Interval <> 2000 Then
    Timer1.Interval = 2000
    Form1.BackColor = RGB(231, 231, 231)
    If Text2.PasswordChar = "*" Then
        Text2.Text = Left("**********************", Len(Text2.Text))
    End If
Else
    If Text2.Text = "wyr" Then
        Pass = 1
        Timer1.Enabled = False
    Else
        Timer1.Enabled = False
    End If
End If
End Sub
