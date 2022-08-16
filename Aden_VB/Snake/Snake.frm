VERSION 5.00
Begin VB.Form Start_Form 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   15240
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Opction_Command 
      Caption         =   "Opction"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5385
   End
   Begin VB.CommandButton Start_Button 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   390
      TabIndex        =   1
      Top             =   2430
      Width           =   5385
   End
   Begin VB.Label The_Snake_Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "The Snake"
      BeginProperty Font 
         Name            =   "AR CENA"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   405
      TabIndex        =   0
      Top             =   525
      Width           =   5265
   End
End
Attribute VB_Name = "Start_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Start_Form.Scale (-10, 10)-(10, -10)
The_Snake_Label.Left = -3
The_Snake_Label.Top = 10
Start_Button.Left = -3
Start_Button.Top = 6
Opction_Command.Left = -3
Opction_Command.Top = 1
End Sub

Private Sub Start_Button_Click()
Game_Form.Show
Start_Form.Hide
End Sub
