VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   513
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1022
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "全部擦除"
      Height          =   375
      Left            =   12675
      TabIndex        =   9
      Top             =   2250
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "改变背景色"
      Height          =   375
      Left            =   11475
      TabIndex        =   8
      Top             =   2250
      Width           =   1125
   End
   Begin VB.HScrollBar RGB1 
      Height          =   300
      Index           =   2
      LargeChange     =   2
      Left            =   11400
      Max             =   256
      TabIndex        =   6
      Top             =   1260
      Width           =   1950
   End
   Begin VB.HScrollBar RGB1 
      Height          =   300
      Index           =   1
      LargeChange     =   2
      Left            =   11400
      Max             =   256
      TabIndex        =   5
      Top             =   885
      Width           =   1950
   End
   Begin VB.HScrollBar RGB1 
      Height          =   300
      Index           =   0
      LargeChange     =   2
      Left            =   11400
      Max             =   256
      TabIndex        =   3
      Top             =   495
      Width           =   1950
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   13320
      TabIndex        =   2
      Top             =   75
      Width           =   1875
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   11385
      TabIndex        =   1
      Top             =   75
      Width           =   1875
   End
   Begin VB.PictureBox Picture1 
      Height          =   7500
      Left            =   45
      MousePointer    =   2  'Cross
      ScaleHeight     =   7440
      ScaleWidth      =   11190
      TabIndex        =   0
      Top             =   90
      Width           =   11250
   End
   Begin VB.Shape CLICK 
      BorderWidth     =   3
      Height          =   300
      Left            =   11490
      Top             =   1695
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Color 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   4
      Left            =   14175
      TabIndex        =   14
      Top             =   1650
      Width           =   600
   End
   Begin VB.Label Color 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   13485
      TabIndex        =   13
      Top             =   1650
      Width           =   600
   End
   Begin VB.Label Color 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   2
      Left            =   12825
      TabIndex        =   12
      Top             =   1650
      Width           =   600
   End
   Begin VB.Label Color 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   12150
      TabIndex        =   11
      Top             =   1650
      Width           =   600
   End
   Begin VB.Label Color 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   11475
      TabIndex        =   10
      Top             =   1650
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "label2"
      Height          =   180
      Left            =   11595
      TabIndex        =   7
      Top             =   2070
      Width           =   540
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   1080
      Left            =   13470
      TabIndex        =   4
      Top             =   480
      Width           =   1560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private TheX As Single
Private TheY As Single
Private step As Integer
Private Press As Boolean
Private Check As Integer

Private Sub Color_Click(Index As Integer)
CLICK.Visible = True
CLICK.Top = Color(Index).Top
CLICK.Left = Color(Index).Left
Check = Index
End Sub

Private Sub Color_DblClick(Index As Integer)
Color(Check).BackColor = Label1.BackColor
End Sub

Private Sub Command1_Click()
Picture1.BackColor = Label1.BackColor
End Sub

Private Sub Command2_Click()
Picture1.Cls
End Sub

Private Sub Form_Load()
Picture1.Scale (0, 0)-(75, 50)
Picture1.Left = 0
Picture1.Top = 0
Label2.Caption = ""
Check = 5
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Point As POINTAPI
GetCursorPos Point
If KeyCode = 32 Then
    Press = True
    Call Draw
End If
If KeyCode >= 48 And KeyCode < 58 Then step = KeyCode - 48
If step = 0 Then step = 10
If KeyCode = 39 Then
    Call Draw: SetCursorPos (Point.X + step), (Point.Y)
ElseIf KeyCode = 37 Then
    Call Draw: SetCursorPos (Point.X + -step), (Point.Y)
ElseIf KeyCode = 38 Then
    Call Draw: SetCursorPos (Point.X), (Point.Y + -step)
ElseIf KeyCode = 40 Then
    Call Draw: SetCursorPos (Point.X), (Point.Y + step)
End If
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then Press = False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    For z = 0 To 1 Step 0.1
        Picture1.Line (Int(X + 0.25) + z, Int(Y + 0.25))-(Int(X + 0.25) + z, Int(Y + 0.25) + 1), Label1.BackColor
    Next z
End If
If Button = 2 And Check <> 5 Then Color(Check).BackColor = Picture1.Point(X, Y)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TheX = X
TheY = Y
Text1.Text = Int(X * 100 + 0.5) / 100
Text2.Text = Int(Y * 100 + 0.5) / 100
End Sub

Private Sub RGB1_Change(Index As Integer)
Label1.BackColor = RGB(RGB1(0).Value, RGB1(1).Value, RGB1(2).Value)
Label2.Caption = "R: " & RGB1(0).Value & " G: " & RGB1(1).Value & " B: " & RGB1(2).Value
End Sub

Private Sub Draw()
If Press = True Then
    For z = 0 To 1 Step 0.1
        Picture1.Line (Int(TheX + 0.25) + z, Int(TheY + 0.25))-(Int(TheX + 0.25) + z, Int(TheY + 0.25) + 1), Label1.BackColor
    Next z
End If
End Sub

