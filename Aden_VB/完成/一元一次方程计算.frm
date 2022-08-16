VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "一元一次方程计算器"
   ClientHeight    =   5985
   ClientLeft      =   2220
   ClientTop       =   3000
   ClientWidth     =   15585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   15585
   Begin VB.CommandButton Command4 
      Caption         =   "一元二次方程计算器"
      Height          =   735
      Left            =   7680
      TabIndex        =   7
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      TabIndex        =   4
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "x="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "X+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4320
      TabIndex        =   2
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "=0"
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
      Left            =   6960
      TabIndex        =   0
      Top             =   1200
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command3_Click()
If Text2 = "" Then
   Text2 = "0"
End If
If Text1 = "" Then
   Text1 = "0"
End If
Text3 = Val(-Text1) / Val(Text2)
End Sub


Private Sub Command4_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Form_Load()
If Date >= #8/15/2016# And Date < #8/30/2016# Then
   X = MsgBox("此应用程序受权期以过，请连系程序编写者：陈锡安", 48, "提示"): Form1.Hide: Form2.Hide: Form3.Show
ElseIf Date >= #8/30/2016# Then
   X = MsgBox("此应用程序受权期以过，请连系程序编写者：陈锡安", 48, "提示"): Form1.Hide: Form2.Hide: Form3.Show
End If
End Sub

Private Sub Text2_Click()
Text1 = ""
Text2 = ""
Text3 = ""
End Sub
'Edited by Aden
'2016/7/28
