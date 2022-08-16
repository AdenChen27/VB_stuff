VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "商店"
   ClientHeight    =   5715
   ClientLeft      =   3735
   ClientTop       =   2145
   ClientWidth     =   9810
   LinkTopic       =   "Form4"
   ScaleHeight     =   5715
   ScaleWidth      =   9810
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   30
      TabIndex        =   4
      Top             =   405
      Visible         =   0   'False
      Width           =   9855
      Begin VB.Frame Frame2 
         Height          =   5055
         Left            =   135
         TabIndex        =   11
         Top             =   45
         Visible         =   0   'False
         Width           =   9855
         Begin VB.CommandButton Command9 
            Caption         =   "十秒"
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
            Left            =   5400
            TabIndex        =   15
            Top             =   2160
            Width           =   3135
         End
         Begin VB.CommandButton Command8 
            Caption         =   "五秒"
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
            Left            =   720
            TabIndex        =   14
            Top             =   2160
            Width           =   3135
         End
         Begin VB.CommandButton Command7 
            Caption         =   "返回"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            TabIndex        =   12
            Top             =   3720
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "使用时，按‘x’"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   20.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   16
            Top             =   840
            Width           =   3165
         End
         Begin VB.Label Label5 
            Caption         =   "穿墙机：可直接穿过墙。"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   20.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   285
            TabIndex        =   13
            Top             =   345
            Width           =   4545
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "十秒"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5280
         TabIndex        =   8
         Top             =   2400
         Width           =   3615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "五秒"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   2400
         Width           =   3615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "反回"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   5
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "使用时，按‘z’"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   20.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   495
         TabIndex        =   9
         Top             =   1170
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   "减速器：让方块的速度慢下来"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   330
         TabIndex        =   6
         Top             =   405
         Width           =   8175
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "减速器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   4920
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "穿墙机"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4800
      TabIndex        =   10
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   9720
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Caption         =   "你的分数："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form4.Hide
Form2.Show
End Sub

Private Sub Command2_Click()
Frame1.Visible = True
End Sub

Private Sub Command3_Click()
Frame1.Visible = False
End Sub

Private Sub Command4_Click()
If s2 >= 5 Then
   s2 = s2 - 5
   Label2.Caption = s2
   sh(0) = sh(0) + 1
   SaveSetting App.Title, "Settings", "s2", s2
ElseIf s2 < 5 Then
   s = MsgBox("not enough points", 48, "")
End If
End Sub

Private Sub Command5_Click()
If s2 >= 10 Then
   s2 = s2 - 10
   Label2.Caption = s2
   sh(0) = sh(0) + 2
   SaveSetting App.Title, "Settings", "s2", s2
ElseIf s2 < 10 Then
   s = MsgBox("not enough points", 48, "")
End If
End Sub

Private Sub Command6_Click()
Frame1.Visible = True
Frame2.Visible = True
End Sub

Private Sub Command7_Click()
Frame2.Visible = False
Frame1.Visible = False
End Sub

Private Sub Command8_Click()
If s2 >= 5 Then
   s2 = s2 - 5
   Label2.Caption = s2
   sh(1) = sh(1) + 1
   SaveSetting App.Title, "Settings", "s2", s2
ElseIf s2 < 5 Then
   s = MsgBox("not enough points", 48, "")
End If
End Sub

Private Sub Command9_Click()
If s2 >= 10 Then
   s2 = s2 - 10
   Label2.Caption = s2
   sh(1) = sh(1) + 2
   SaveSetting App.Title, "Settings", "s2", s2
ElseIf s2 < 10 Then
   s = MsgBox("分数不足", 48, "")
End If
End Sub

Private Sub Form_Activate()
Label2.Caption = s2
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.Title, "Settings", "sh(0)", sh(0)
SaveSetting App.Title, "Settings", "sh(1)", sh(1)
SaveSetting App.Title, "Settings", "s2", s2
Unload Form1
Unload Form2
Unload Form3
End Sub

Private Sub Label1_Click()
s2 = s2 + 10
End Sub

'**********************************************END******************************************
'Edited by Aden
'2016/8/21

