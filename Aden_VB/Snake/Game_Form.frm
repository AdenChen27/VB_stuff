VERSION 5.00
Begin VB.Form Game_Form 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game_Form"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   16
   ScaleMode       =   0  'User
   ScaleWidth      =   32
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Snake_Walk 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Food 
      Height          =   1005
      Left            =   2085
      Top             =   2430
      Width           =   990
   End
   Begin VB.Image Snake_Body 
      Height          =   495
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Snake_Body 
      Height          =   495
      Index           =   0
      Left            =   330
      Top             =   255
      Width           =   1215
   End
End
Attribute VB_Name = "Game_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Last_Direction = Direction
If KeyCode = 38 Then
    Direction = "Up"
ElseIf KeyCode = 40 Then
    Direction = "Down"
ElseIf KeyCode = 37 Then
    Direction = "Left"
ElseIf KeyCode = 39 Then
    Direction = "Right"
End If
End Sub

Private Sub Form_Load()
Snake_Body(0).Left = 1
Snake_Body(0).Top = 0
Snake_Body(0).Picture = LoadPicture("C:\Users\zhchen\Desktop\Aden_VB\Snake\Snake.jpg")
Snake_Body(1).Left = 0
Snake_Body(1).Top = 0
Snake_Body(1).Picture = LoadPicture("C:\Users\zhchen\Desktop\Aden_VB\Snake\Snake.jpg")
Direction = "Right"
Last_Direction = "Right"
Snake_Lenth = 1
Snake_Walk.Enabled = True

Food.Picture = LoadPicture("C:\Users\zhchen\Desktop\Aden_VB\Snake\Food.jpg")
Randomize
Food.Top = Int(Rnd * 16)
Food.Left = Int(Rnd * 32)
End Sub

Private Sub Snake_Walk_Timer()
If Direction = "Right" Then
    If Last_Direction <> "Left" Then
        Snake_Body(0).Left = Snake_Body(0).Left + 1
    Else:
        Direction = "Left"
        Snake_Body(0).Left = Snake_Body(0).Left - 1
    End If
ElseIf Direction = "Left" Then
    If Last_Direction <> "Right" Then
        Snake_Body(0).Left = Snake_Body(0).Left - 1
    Else:
        Direction = "Right"
        Snake_Body(0).Left = Snake_Body(0).Left + 1
    End If
ElseIf Direction = "Up" Then
    If Last_Direction <> "Down" Then
        Snake_Body(0).Top = Snake_Body(0).Top - 1
    Else:
        Direction = "Down"
        Snake_Body(0).Top = Snake_Body(0).Top + 1
    End If
ElseIf Direction = "Down" Then
    If Last_Direction <> "Up" Then
        Snake_Body(0).Top = Snake_Body(0).Top + 1
    Else:
        Direction = "Up"
        Snake_Body(0).Top = Snake_Body(0).Top - 1
    End If
End If
If Pos_X(0) - Food.Left < 1 And Pos_Y(0) - Food.Top < 1 Then
    Load Snake_Body(Snake_Lenth + 1)
    Snake_Body(Snake_Lenth + 1).Left = Pos_X(Snake_Lenth)
    Snake_Body(Snake_Lenth + 1).Top = Pos_Y(Snake_Lenth)
    Snake_Body(Snake_Lenth + 1).Visible = True
    Snake_Lenth = Snake_Lenth + 1
    Food.Top = Int(Rnd * 16)
    Food.Left = Int(Rnd * 32)
End If
For i = Snake_Lenth To 1 Step -1
    Pos_X(i) = Pos_X(i - 1)
    Pos_Y(i) = Pos_Y(i - 1)
    Snake_Body(i).Left = Pos_X(i)
    Snake_Body(i).Top = Pos_Y(i)
    Pos_X(0) = Snake_Body(0).Left
    Pos_Y(0) = Snake_Body(0).Top
Next i
End Sub
