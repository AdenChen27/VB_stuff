VERSION 5.00
Begin VB.Form Color_Form 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "颜色"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   495
      TabIndex        =   5
      Top             =   5505
      Width           =   2805
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2580
      Index           =   4
      Left            =   5520
      ScaleHeight     =   2520
      ScaleWidth      =   2520
      TabIndex        =   4
      Top             =   2805
      Width           =   2580
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2580
      Index           =   3
      Left            =   2880
      ScaleHeight     =   2520
      ScaleWidth      =   2520
      TabIndex        =   3
      Top             =   2820
      Width           =   2580
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2580
      Index           =   2
      Left            =   5505
      ScaleHeight     =   2520
      ScaleWidth      =   2520
      TabIndex        =   2
      Top             =   195
      Width           =   2580
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2580
      Index           =   1
      Left            =   2865
      ScaleHeight     =   2520
      ScaleWidth      =   2520
      TabIndex        =   1
      Top             =   180
      Width           =   2580
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2580
      Index           =   0
      Left            =   195
      ScaleHeight     =   2520
      ScaleWidth      =   2520
      TabIndex        =   0
      Top             =   180
      Width           =   2580
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   180
      Top             =   2850
      Width           =   2565
   End
End
Attribute VB_Name = "Color_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Color_That_Change = 0 Then Form2.Color_Squre.BackColor = Shape1.FillColor Else: Form2.Color_Obstacle.BackColor = Shape1.FillColor
End Sub

Private Sub Form_Load()
For X = 0 To 256 Step 3
    For Y = 0 To 256 Step 3
        For z = 0 To 4
            Picture1(z).PSet (X * 10, Y * 10), RGB(X, Y, z * 50)
            Picture1(z).PSet (X * 10 + 15, Y * 10), RGB(X, Y, z * 50)
            Picture1(z).PSet (X * 10, Y * 10 + 15), RGB(X, Y, z * 50)
            Picture1(z).PSet (X * 10 + 15, Y * 10 + 15), RGB(X, Y, z * 50)
        Next z
    Next Y
Next X
End Sub

Private Sub Picture1_Mousedown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.FillColor = RGB(X / 10, Y / 10, Index * 50)
Text2 = Str(Int(X / 10)) + "," + Str(Int(Y / 10)) + "," + Str(Int(Index * 50))
End Sub
