VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "输入密码"
   ClientHeight    =   2505
   ClientLeft      =   7845
   ClientTop       =   4575
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   20.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   2505
   ScaleWidth      =   5115
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "密码："
      Height          =   405
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Date >= #8/15/2016# And Date < #8/30/2016# And Text1 = "aden8051" Then
   Form3.Hide: Form1.Show
ElseIf Date >= #8/15/2016# And Date < #8/30/2016# Then
   y = MsgBox("密码错误", 48, "提示")
End If
If Date >= #8/30/2016# And Text1 = "aden8003" Then
   Form3.Hide: Form1.Show
ElseIf Date >= #8/30/2016# Then
   y = MsgBox("密码错误", 48, "提示")
End If
End Sub
'Edited by Aden
'2016/8/1
