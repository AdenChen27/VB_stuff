VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "计算器"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   150
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   606
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   1590
      TabIndex        =   7
      Top             =   75
      Visible         =   0   'False
      Width           =   7440
   End
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   1650
      TabIndex        =   3
      Top             =   75
      Width           =   7230
      Begin VB.TextBox Ans 
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
         Left            =   270
         TabIndex        =   6
         Top             =   1020
         Width           =   4830
      End
      Begin VB.CommandButton Calculate 
         Caption         =   "计算"
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
         Left            =   5190
         TabIndex        =   5
         Top             =   1020
         Width           =   1830
      End
      Begin VB.TextBox NumText 
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
         Left            =   255
         TabIndex        =   4
         Top             =   285
         Width           =   6735
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "分数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1275
      Width           =   1275
   End
   Begin VB.OptionButton WithX 
      Caption         =   "方程"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   750
      Width           =   1275
   End
   Begin VB.OptionButton Normal 
      Caption         =   "普通"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Num(50) As Double
Private Symbol(50) As String
Private NumOfNum As Integer
Private x As Long

Private Sub Calculate_Click()
Dim FirstX As Integer
FirstX = 1
Times = 0
NumOfNum = 0
For x = 0 To 50
    Num(x) = 0
    Symbol(x) = ""
Next x

NumText.Text = NumText.Text & "+1 "
For x = 1 To Len(NumText.Text)
    If Asc(Mid(NumText.Text, x, 1)) = 42 Or Asc(Mid(NumText.Text, x, 1)) = 43 Or _
       Asc(Mid(NumText.Text, x, 1)) = 45 Or Asc(Mid(NumText.Text, x, 1)) = 47 Then
        Num(NumOfNum) = Val(Mid(NumText, FirstX, x - FirstX))
        FirstX = x + 1
        NumOfNum = NumOfNum + 1
    ElseIf Asc(Mid(NumText.Text, x, 1)) = 32 Then
        Num(NumOfNum) = Val(Mid(NumText, Len(NumText) - FirstX))
        NumOfNum = NumOfNum + 1
    End If
Next x
NumText.Text = Left(NumText.Text, Len(NumText.Text) - 3)

NumOfNum = 0
For x = 1 To Len(NumText.Text)
    If Asc(Mid(NumText.Text, x, 1)) = 42 Or Asc(Mid(NumText.Text, x, 1)) = 43 Or _
       Asc(Mid(NumText.Text, x, 1)) = 45 Or Asc(Mid(NumText.Text, x, 1)) = 47 Then
         Symbol(NumOfNum) = Mid(NumText.Text, x, 1)
         NumOfNum = NumOfNum + 1
    End If
Next x

c = 0
For x = 0 To 50
    If Symbol(x) = "-" Then a = a + 1
    If Symbol(x) = "/" Then b = b + 1
    If b > 2 Or a > 2 Then a = MsgBox("项目不匹配!", vbExclamation, "提示"): Exit For: c = 1
Next x
If c = 0 Then
For a = 0 To NumOfNum - 1
For x = 0 To NumOfNum - 1
    If Symbol(x) = "*" Then
        Num(x) = Num(x) * Num(x + 1)
        Call MoveTheNum
        x = 0
    ElseIf Symbol(x) = "/" Then
        Num(x) = Num(x) / Num(x + 1)
        Call MoveTheNum
        x = 0
    End If
Next x
Next a
For a = 0 To NumOfNum - 1
For x = 0 To NumOfNum - 1
    If Symbol(x) = "+" Then
        Num(x) = Num(x) + Num(x + 1)
        Call MoveTheNum
        x = 0
    ElseIf Symbol(x) = "-" Then
        Num(x) = Num(x) - Num(x + 1)
        Call MoveTheNum
        x = 0
    End If
Next x
Next a
Ans = Num(0)
End If
End Sub

Private Sub Normal_Click()
Frame1.Visible = True
Frame2.Visible = False
End Sub

Private Sub NumText_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 42 Or KeyAscii = 45 Or KeyAscii = 47 _
   Or KeyAscii = 43 Or KeyAscii = 46 Or (KeyAscii >= 48 And KeyAscii <= 57) Then Exit Sub Else: KeyAscii = 0
End Sub

Private Sub MoveTheNum()
For Y = x + 1 To 49
    Num(Y) = Num(Y + 1)
    Symbol(Y - 1) = Symbol(Y)
Next Y
End Sub

Private Sub WithX_Click()
Frame2.Visible = True
Frame1.Visible = False
End Sub
