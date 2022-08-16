VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "颜色"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   11415
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   2910
      TabIndex        =   13
      Top             =   5475
      Width           =   2295
   End
   Begin VB.VScrollBar blue 
      Height          =   2310
      LargeChange     =   10
      Left            =   10500
      Max             =   256
      SmallChange     =   5
      TabIndex        =   11
      Top             =   3225
      Width           =   330
   End
   Begin VB.VScrollBar green 
      Height          =   2310
      LargeChange     =   10
      Left            =   9645
      Max             =   256
      SmallChange     =   5
      TabIndex        =   10
      Top             =   3225
      Width           =   330
   End
   Begin VB.VScrollBar red 
      Height          =   2310
      LargeChange     =   10
      Left            =   8685
      Max             =   256
      SmallChange     =   5
      TabIndex        =   9
      Top             =   3225
      Width           =   330
   End
   Begin VB.PictureBox Picture2 
      Height          =   2580
      Left            =   8430
      ScaleHeight     =   2520
      ScaleWidth      =   2520
      TabIndex        =   7
      Top             =   15
      Width           =   2580
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   1230
      TabIndex        =   6
      Top             =   5460
      Width           =   1560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "渐变"
      Height          =   405
      Left            =   180
      TabIndex        =   5
      Top             =   5475
      Width           =   840
   End
   Begin VB.PictureBox Picture1 
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
      Height          =   2580
      Index           =   0
      Left            =   195
      ScaleHeight     =   2520
      ScaleWidth      =   2520
      TabIndex        =   0
      Top             =   180
      Width           =   2580
   End
   Begin VB.Label Label2 
      Height          =   390
      Left            =   8505
      TabIndex        =   12
      Top             =   5625
      Width           =   2700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "R   G   B"
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
      Left            =   8745
      TabIndex        =   8
      Top             =   2700
      Width           =   2025
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   180
      Top             =   2850
      Width           =   2565
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub blue_Change()
Picture2.BackColor = RGB(red.Value, green.Value, blue.Value)
Label2.Caption = "R: " & red.Value & " G: " & green.Value & " B: " & blue.Value
End Sub

Private Sub Command2_Click()
Picture1(0).Cls
Picture1(1).Cls
Picture1(2).Cls
Picture1(3).Cls
Picture1(4).Cls
If Val(Text1.Text) > 0.1 Then
    For X = 0 To 256 Step Val(Text1.Text)
        For Y = 0 To 256 Step Val(Text1.Text)
            For z = 0 To 4
                Picture1(z).PSet (X * 10, Y * 10), RGB(X, Y, z * 50)
                Picture1(z).PSet (X * 10 + 15, Y * 10), RGB(X, Y, z * 50)
                Picture1(z).PSet (X * 10, Y * 10 + 15), RGB(X, Y, z * 50)
                Picture1(z).PSet (X * 10 + 15, Y * 10 + 15), RGB(X, Y, z * 50)
            Next z
        Next Y
    Next X
Else
    a = MsgBox("无效步长值", , "")
End If
End Sub

Private Sub green_Change()
Picture2.BackColor = RGB(red.Value, green.Value, blue.Value)
Label2.Caption = "R: " & red.Value & " G: " & green.Value & " B: " & blue.Value
End Sub

Private Sub Picture1_Mousedown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.FillColor = RGB(X / 10, Y / 10, Index * 50)
Text2 = Str(Int(X / 10)) + "," + Str(Int(Y / 10)) + "," + Str(Int(Index * 50))
End Sub

Private Sub red_Change()
Picture2.BackColor = RGB(red.Value, green.Value, blue.Value)
Label2.Caption = "R: " & red.Value & " G: " & green.Value & " B: " & blue.Value
End Sub
