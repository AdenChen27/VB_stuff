VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "一元二次方程计算器"
   ClientHeight    =   5715
   ClientLeft      =   2220
   ClientTop       =   3135
   ClientWidth     =   15435
   LinkTopic       =   "Form2"
   ScaleHeight     =   5715
   ScaleWidth      =   15435
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
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1215
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
      TabIndex        =   9
      Top             =   120
      Width           =   1215
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
      Left            =   4560
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text4 
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
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text5 
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
      Left            =   3720
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
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
      Left            =   600
      TabIndex        =   5
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "一元一次方程计算器"
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "作图"
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
      Left            =   8880
      TabIndex        =   3
      Top             =   5040
      Width           =   1335
   End
   Begin VB.PictureBox P1 
      Height          =   4935
      Left            =   6840
      ScaleHeight     =   4875
      ScaleWidth      =   8595
      TabIndex        =   2
      Top             =   0
      Width           =   8655
   End
   Begin VB.TextBox Text6 
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
      Left            =   1080
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清除"
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
      Left            =   10800
      TabIndex        =   0
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   17
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "X^2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   18
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "X"
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
      Left            =   3720
      TabIndex        =   16
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "+"
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
      Left            =   4080
      TabIndex        =   15
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "=0"
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
      Left            =   6000
      TabIndex        =   14
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "X1="
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
      Left            =   0
      TabIndex        =   13
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "X2="
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
      Left            =   2760
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Δ="
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
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1 = "" Then
   Text1 = "0"
End If
If Text2 = "" Then
   Text2 = "0"
End If
If Text3 = "" Then
   Text3 = "0"
End If
Text6 = Text2 ^ 2 - 4 * Text1 * Text3
If Text6 < 0 Then
   X = MsgBox("此方程无解", 48, "提示")
Else
   Text4 = (Val(-Text2) + Sqr(Text6)) / 2 * Text1
   Text5 = (Val(-Text2) - Sqr(Text6)) / 2 * Text1
End If
End Sub


Private Sub Command3_Click()
P1.Picture = LoadPicture("")
End Sub

Private Sub Form_Load()
P1.Scale (-10, 10)-(10, -10)
End Sub

Private Sub Command2_Click()
If Text1 = "" Then
   Text1 = "0"
End If
If Text2 = "" Then
   Text2 = "0"
End If
If Text3 = "" Then
   Text3 = "0"
End If
Dim X As Single
P1.Line (-10, 0)-(10, 0), vbRed
P1.Line (0, -10)-(0, 10), vbRed
For X = -10 To 10 Step 0.02
P1.PSet (X, Text1 * X ^ 2 + Text2 * X + Text3)
Next X
End Sub


Private Sub Command4_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Text1_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
End Sub
'Edited by Aden
'2016/7/28

