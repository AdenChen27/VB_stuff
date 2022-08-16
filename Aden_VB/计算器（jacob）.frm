VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12885
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   26.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   12885
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   4410
      TabIndex        =   9
      Top             =   5040
      Width           =   3060
   End
   Begin VB.CommandButton command4 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   9810
      TabIndex        =   8
      Top             =   3465
      Width           =   2235
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   8400
      TabIndex        =   5
      Top             =   1050
      Width           =   2415
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
      Height          =   825
      Left            =   4575
      TabIndex        =   4
      Top             =   1065
      Width           =   2415
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
      Height          =   885
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   300
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   3585
      TabIndex        =   1
      Top             =   3435
      Width           =   2130
   End
   Begin VB.CommandButton Command3 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6645
      TabIndex        =   0
      Top             =   3480
      Width           =   2190
   End
   Begin VB.Label Label2 
      Caption         =   " ="
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
      Left            =   7005
      TabIndex        =   7
      Top             =   1110
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      Left            =   3240
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
Label1 = "+"
End Sub

Private Sub Command2_Click()
Label1 = "-"
End Sub

Private Sub Command3_Click()
Label1 = "*"
End Sub

Private Sub command4_Click()
Label1 = "/"
End Sub

Private Sub Command5_Click()
If Label1 = "+" Then Text3 = Val(Text1) + Val(Text2)
If Label1 = "-" Then Text3 = Val(Text1) - Val(Text2)
If Label1 = "*" Then Text3 = Val(Text1) * Val(Text2)
If Label1 = "/" Then Text3 = Val(Text1) / Val(Text2)
End Sub


