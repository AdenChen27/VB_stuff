VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2745
   ClientLeft      =   6975
   ClientTop       =   3990
   ClientWidth     =   5310
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
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
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "密码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Date >= #9/15/2016# And Date < #9/30/2016# And Text1 = "aden8051" Then
   Form2.Hide: Form1.Show
ElseIf Date >= #9/15/2016# And Date < #9/30/2016# Then
   Y = MsgBox("密码错误", 48, "提示")
End If
If Date >= #9/30/2016# And Text1 = "aden8003" Then
   Form2.Hide: Form1.Show
ElseIf Date >= #9/30/2016# Then
   Y = MsgBox("密码错误", 48, "提示")
End If
End Sub
'Edited by Aden
'2016/8/7

Private Sub Form_Unload(Cancel As Integer)
Unload Form1
Unload Form3
End Sub
