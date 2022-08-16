VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "打字游戏"
   ClientHeight    =   5160
   ClientLeft      =   3810
   ClientTop       =   3405
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   12360
   Begin VB.Frame Frame1 
      Height          =   4980
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   12285
      Begin VB.CommandButton Command1 
         Caption         =   "确定"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4020
         TabIndex        =   7
         Top             =   2250
         Width           =   3855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "规则：按照屏幕键盘上方的字母打字。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1785
         TabIndex        =   6
         Top             =   885
         Width           =   7395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "打字游戏"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5070
         TabIndex        =   5
         Top             =   300
         Width           =   1740
      End
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "正确率：0%"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   2385
      TabIndex        =   8
      Top             =   240
      Width           =   5340
   End
   Begin VB.Label key3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " z"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1000
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   3600
      Width           =   1000
   End
   Begin VB.Label key2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " a"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1000
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   1000
   End
   Begin VB.Label key 
      BackColor       =   &H00E0E0E0&
      Caption         =   " q"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1000
      Index           =   0
      Left            =   225
      TabIndex        =   1
      Top             =   1200
      Width           =   1000
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "字母："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   255
      TabIndex        =   0
      Top             =   195
      Width           =   1980
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Correct As Integer
Private Wrong As Integer
Public Sub TimeDelay(ByVal PauseSecond As Single)
 Dim Star, PauseTime
 Star = Timer
 PauseTime = PauseSecond
 Do While Timer < Star + PauseTime
 DoEvents
 Loop
End Sub

Private Sub Command1_Click()
Frame1.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If LCase(Chr(KeyCode)) = Right(Label1.Caption, 1) Then
    Label1.Caption = "字母：" + Chr(Int((97 - 122 + 1) * Rnd() + 122))
    Correct = Correct + 1
    Label4.Caption = "正确率：" + Str(Int(Correct / (Correct + Wrong) * 10000) / 100) + "%"
    For x = 0 To 9
        If Right(key(x).Caption, 1) = LCase(Chr(KeyCode)) Then key(x).BackColor = &HFF00&: TimeDelay (0.2): key(x).BackColor = &HE0E0E0: Exit For
    Next x
    For x = 0 To 8
        If Right(key2(x).Caption, 1) = LCase(Chr(KeyCode)) Then key2(x).BackColor = &HFF00&: TimeDelay (0.2): key2(x).BackColor = &HE0E0E0: Exit For
    Next x
    For x = 0 To 6
        If Right(key3(x).Caption, 1) = LCase(Chr(KeyCode)) Then key3(x).BackColor = &HFF00&: TimeDelay (0.2): key3(x).BackColor = &HE0E0E0: Exit For
    Next x
Else
    Wrong = Wrong + 1
    Label4.Caption = "正确率：" + Str(Int(Correct / (Correct + Wrong) * 10000) / 100) + "%"
    For x = 0 To 9
        If Right(key(x).Caption, 1) = LCase(Chr(KeyCode)) Then key(x).BackColor = &HFF&: TimeDelay (0.2): key(x).BackColor = &HE0E0E0: Exit For
    Next x
    For x = 0 To 8
        If Right(key2(x).Caption, 1) = LCase(Chr(KeyCode)) Then key2(x).BackColor = &HFF&: TimeDelay (0.2): key2(x).BackColor = &HE0E0E0: Exit For
    Next x
    For x = 0 To 6
        If Right(key3(x).Caption, 1) = LCase(Chr(KeyCode)) Then key3(x).BackColor = &HFF&: TimeDelay (0.2): key3(x).BackColor = &HE0E0E0: Exit For
    Next x
End If
End Sub

Private Sub Form_Load()
letter = "wertyuiop"
Randomize
For x = 1 To 9
    Load key(x)
    key(x).Visible = True
    key(x).Left = key(x - 1).Left + 1200
    key(x).Top = 1200
    key(x).Caption = " " + Left(letter, 1)
    letter = Right(letter, Len(letter) - 1)
Next x
letter = "sdfghjkl"
For x = 1 To 8
    Load key2(x)
    key2(x).Visible = True
    key2(x).Left = key2(x - 1).Left + 1200
    key2(x).Top = 2400
    key2(x).Caption = " " + Left(letter, 1)
    letter = Right(letter, Len(letter) - 1)
Next x
letter = "xcvbnm"
For x = 1 To 6
    Load key3(x)
    key3(x).Visible = True
    key3(x).Left = key3(x - 1).Left + 1200
    key3(x).Top = 3600
    key3(x).Caption = " " + Left(letter, 1)
    letter = Right(letter, Len(letter) - 1)
Next x
Label1.Caption = "字母：" + Chr(Int((97 - 122 + 1) * Rnd() + 122))
End Sub
