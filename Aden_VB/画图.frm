VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   17190
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CodeIn 
      Height          =   1800
      Left            =   15690
      TabIndex        =   4
      Top             =   4380
      Width           =   1275
   End
   Begin VB.TextBox CodeText 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   7425
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   4335
      Width           =   8250
   End
   Begin VB.TextBox Code 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3995
      Left            =   7395
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   105
      Width           =   9500
   End
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      Height          =   15000
      Left            =   30
      MousePointer    =   2  'Cross
      ScaleHeight     =   14940
      ScaleWidth      =   6840
      TabIndex        =   0
      Top             =   30
      Width           =   6900
   End
   Begin VB.Label CodeTell 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7485
      TabIndex        =   5
      Top             =   6345
      Width           =   8430
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   17.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7980
      TabIndex        =   1
      Top             =   1380
      Width           =   2700
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

Private Sub CodeIn_Click()
If Left(CodeText, 6) = "shape4" Then
    Call FindN(4, 8)
    P1.Line (n(1), n(2))-(n(1) + n(3), n(2))
    P1.Line (n(1), n(2))-(n(1), n(2) - wi)
    P1.Line (n(1) + n(3), n(2))-(n(1) + n(3), n(2) - n(4))
    P1.Line (n(1), n(2) - n(4))-(n(1) + n(3), n(2) - n(4))
    Code = Code & CodeText & vbCrLf
Else
    a = MsgBox("此函数不存在", vbInformation)
    CodeText = ""
End If
End Sub

Private Sub CodeText_Change()
If Left(CodeText, 6) = "shape4" Then
    CodeTell.Caption = "shape4(x,y,long,width)"
Else
    CodeTell = ""
End If
End Sub

Private Sub Form_Load()
Form1.Height = Screen.Height
Form1.Width = Screen.Width
Form1.Top = 0
Form1.Left = 0
P1.Height = 10500
P1.Width = 10500
P1.Scale (-1000, 1000)-(1000, -1000)
Label1.Left = 50
Label1.Top = 10500
Code.Left = 10600
CodeText.Left = 10600
CodeIn.Left = 19000
CodeTell.Left = 10600
P1.Line (0, -1000)-(0, 1000), vbRed
P1.Line (-1000, 0)-(1000, 0), vbRed
End Sub

Private Sub P1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Point As POINTAPI
GetCursorPos Point
If KeyCode = 39 Then
    SetCursorPos (Point.X + 1), (Point.Y)
ElseIf KeyCode = 37 Then
    SetCursorPos (Point.X + -1), (Point.Y)
ElseIf KeyCode = 38 Then
    SetCursorPos (Point.X), (Point.Y + -1)
ElseIf KeyCode = 40 Then
    SetCursorPos (Point.X), (Point.Y + 1)
End If
End Sub

Private Sub P1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "X:" & Int(X + 0.5) & " " & "Y:" & Int(Y + 0.5)
End Sub
Private Sub FindN(num_of_n, long_of_start)
If num_of_n >= 1 Then
    For X = long_of_start To Len(CodeText)
        If Mid(CodeText, X, 1) = "," Then Exit For
    Next X
    n(1) = Val(Mid(CodeText, long_of_start, X - long_of_start))
    If num_of_n >= 2 Then
        For Y = X + 1 To Len(CodeText)
            If Mid(CodeText, Y, 1) = "," Then Exit For
        Next Y
        n(2) = Val(Mid(CodeText, X + 1, Y - X - 1))
        If num_of_n >= 3 Then
            For l = Y + 1 To Len(CodeText)
                If Mid(CodeText, l, 1) = "," Then Exit For
            Next l
            n(3) = Val(Mid(CodeText, Y + 1, l - Y - 1))
            If num_of_n = 4 Then
                For w = l + 1 To Len(CodeText)
                    If Mid(CodeText, w, 1) = ")" Then Exit For
                Next w
                n(4) = Val(Mid(CodeText, l + 1, w - l - 1))
            End If
        End If
    End If
End If
End Sub

