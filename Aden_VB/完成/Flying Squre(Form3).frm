VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5865
   ClientLeft      =   3645
   ClientTop       =   2055
   ClientWidth     =   10005
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10005
   Begin VB.CommandButton Command3 
      Caption         =   "商店"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   4
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   3
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1005
      TabIndex        =   2
      Top             =   2430
      Width           =   6600
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1275
      TabIndex        =   1
      Top             =   810
      Width           =   6600
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Label1 = ""
Form1.Timer1.Enabled = False
Form1.Timer2.Enabled = False
Form1.Timer3.Enabled = False
Form1.Timer5.Enabled = True
a4 = 0
Form1.Label1.Visible = True
Form1.Shape3.Visible = False
Form1.Show
Form3.Hide
End Sub


Private Sub Command2_Click()
Form2.Show
Form2.Frame1.Visible = True
Form3.Hide
End Sub


Private Sub Command3_Click()
Form3.Hide
Form4.Show
End Sub

Private Sub Form_Activate()
s2 = s2 + s1
Form3.Label1 = "你这一次的分数是 : " & s1
Form3.Label2 = "你的总分数是 :" & s2
SaveSetting App.Title, "Settings", "s2", s2
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.Title, "Settings", "sh(0)", sh(0)
SaveSetting App.Title, "Settings", "s2", s2
SaveSetting App.Title, "Settings", "sh(1)", sh(1)
Unload Form1
Unload Form2
Unload Form4
End Sub

