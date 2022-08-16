VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5895
   ClientLeft      =   3660
   ClientTop       =   2055
   ClientWidth     =   9945
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9945
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VB.OptionButton Option4 
         Caption         =   "很难"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6300
         TabIndex        =   11
         Top             =   3165
         Width           =   1665
      End
      Begin VB.OptionButton Option2 
         Caption         =   "中等"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   10
         Top             =   3120
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         Caption         =   "难"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   9
         Top             =   3120
         Width           =   1365
      End
      Begin VB.OptionButton Option1 
         Caption         =   "容易"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
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
         Left            =   1320
         TabIndex        =   7
         Top             =   5040
         Width           =   1455
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   20.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         ItemData        =   "test 2.frx":0000
         Left            =   2460
         List            =   "test 2.frx":0013
         TabIndex        =   6
         Top             =   1125
         Width           =   2895
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   20.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         ItemData        =   "test 2.frx":0035
         Left            =   2460
         List            =   "test 2.frx":0048
         TabIndex        =   3
         Top             =   180
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "等级："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   27.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   375
         TabIndex        =   13
         Top             =   2295
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "障碍颜色："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   300
         TabIndex        =   5
         Top             =   1260
         Width           =   2400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "方块颜色："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   225
         Width           =   2400
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   9615
      Begin VB.CommandButton Command5 
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
         Height          =   1095
         Left            =   3120
         TabIndex        =   17
         Top             =   3960
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "规则：控制方块避开障碍，接触到小方块时会有一定几率加分。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   16
         Top             =   1200
         Width           =   7455
      End
      Begin VB.Label Label4 
         Caption         =   "作者：陈锡安 "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   8055
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "设置"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "商店"
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
      Left            =   3120
      TabIndex        =   12
      Top             =   3240
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
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
Form2.Hide
End Sub

Private Sub Command2_Click()
Frame1.Visible = True
End Sub

Private Sub Command3_Click()
If List1.ListIndex = 0 Then a2 = 0
If List1.ListIndex = 1 Then a2 = 1
If List1.ListIndex = 2 Then a2 = 2
If List1.ListIndex = 3 Then a2 = 3
If List1.ListIndex = 4 Then a2 = 4
If List2.ListIndex = 0 Then a3 = 0
If List2.ListIndex = 1 Then a3 = 1
If List2.ListIndex = 2 Then a3 = 2
If List2.ListIndex = 3 Then a3 = 3
If List2.ListIndex = 4 Then a3 = 4
If Option1.Value = True Then a5 = 0
If Option2.Value = True Then a5 = 1
If Option3.Value = True Then a5 = 2
If Option4.Value = True Then a5 = 3
Frame1.Visible = False
End Sub


Private Sub Command4_Click()
Form2.Hide
Form4.Show
End Sub

Private Sub Command5_Click()
Frame2.Visible = False
End Sub

Private Sub Form_Load()
a2 = 0
a3 = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.Title, "Settings", "sh(0)", sh(0)
SaveSetting App.Title, "Settings", "s2", s2
SaveSetting App.Title, "Settings", "sh(1)", sh(1)
Unload Form1
Unload Form3
Unload Form4
End Sub
'Edited by Aden
'2016/8/21

