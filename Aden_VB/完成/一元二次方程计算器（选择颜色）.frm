VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8010
   LinkTopic       =   "Form3"
   ScaleHeight     =   5670
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2580
      Index           =   0
      Left            =   15
      ScaleHeight     =   2520
      ScaleWidth      =   2520
      TabIndex        =   5
      Top             =   0
      Width           =   2580
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2580
      Index           =   1
      Left            =   2685
      ScaleHeight     =   2520
      ScaleWidth      =   2520
      TabIndex        =   4
      Top             =   0
      Width           =   2580
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2580
      Index           =   2
      Left            =   5325
      ScaleHeight     =   2520
      ScaleWidth      =   2520
      TabIndex        =   3
      Top             =   15
      Width           =   2580
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2580
      Index           =   3
      Left            =   2700
      ScaleHeight     =   2520
      ScaleWidth      =   2520
      TabIndex        =   2
      Top             =   2640
      Width           =   2580
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2580
      Index           =   4
      Left            =   5340
      ScaleHeight     =   2520
      ScaleWidth      =   2520
      TabIndex        =   1
      Top             =   2625
      Width           =   2580
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   5250
      Width           =   2610
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   0
      Top             =   2670
      Width           =   2565
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
co = Shape1.FillColor
Form3.Hide
End Sub

Private Sub Form_load()
For X = 0 To 256 Step 1.6
    For Y = 0 To 256 Step 1.6
        For z = 0 To 4
            Picture1(z).PSet (X * 10, Y * 10), RGB(X, Y, z * 50)
        Next z
    Next Y
Next X
End Sub

Private Sub Picture1_Mousedown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.FillColor = RGB(X / 10, Y / 10, Index * 50)
End Sub

