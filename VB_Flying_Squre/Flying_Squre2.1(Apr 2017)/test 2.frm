VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5895
   ClientLeft      =   3660
   ClientTop       =   2055
   ClientWidth     =   9945
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9945
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   150
      TabIndex        =   2
      Top             =   45
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CommandButton Command6 
         Caption         =   "确定"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7500
         TabIndex        =   25
         Top             =   2640
         Width           =   1110
      End
      Begin VB.CheckBox Check1 
         Caption         =   "在硬盘中储存游戏信息(默认路径：D:\Flying_Squre)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   135
         TabIndex        =   24
         Top             =   2265
         Width           =   7305
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2805
         TabIndex        =   23
         Top             =   2655
         Width           =   4650
      End
      Begin VB.CommandButton More_Color 
         Caption         =   "更多"
         Height          =   400
         Index           =   1
         Left            =   3500
         TabIndex        =   21
         Top             =   1590
         Width           =   1005
      End
      Begin VB.CommandButton More_Color 
         Caption         =   "更多"
         Height          =   400
         Index           =   0
         Left            =   3500
         TabIndex        =   20
         Top             =   645
         Width           =   1005
      End
      Begin VB.OptionButton Grade_Very_Hard 
         Caption         =   "很难"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5700
         TabIndex        =   11
         Top             =   1050
         Width           =   1065
      End
      Begin VB.OptionButton Grade_Mid 
         Caption         =   "中等"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5700
         TabIndex        =   10
         Top             =   650
         Width           =   960
      End
      Begin VB.OptionButton Grade_Hard 
         Caption         =   "难"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4700
         TabIndex        =   9
         Top             =   1050
         Width           =   1365
      End
      Begin VB.OptionButton Grade_Easy 
         Caption         =   "容易"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4700
         TabIndex        =   8
         Top             =   650
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.CommandButton Command3 
         Caption         =   "返回"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         TabIndex        =   7
         Top             =   5040
         Width           =   1455
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         ItemData        =   "test 2.frx":0000
         Left            =   2000
         List            =   "test 2.frx":0013
         TabIndex        =   6
         Top             =   1125
         Width           =   1500
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         ItemData        =   "test 2.frx":0035
         Left            =   2000
         List            =   "test 2.frx":0048
         TabIndex        =   3
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "更改游戏音乐路径："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   22
         Top             =   2715
         Width           =   2565
      End
      Begin VB.Label Color_Obstacle 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   3500
         TabIndex        =   19
         Top             =   1150
         Width           =   1000
      End
      Begin VB.Label Color_Squre 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   3500
         TabIndex        =   18
         Top             =   195
         Width           =   1000
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   0
         X2              =   9700
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label Speed 
         AutoSize        =   -1  'True
         Caption         =   "难度："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4740
         TabIndex        =   13
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Obstacle_Color 
         AutoSize        =   -1  'True
         Caption         =   "障碍颜色："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   250
         TabIndex        =   5
         Top             =   1260
         Width           =   1800
      End
      Begin VB.Label Squre_Color 
         AutoSize        =   -1  'True
         Caption         =   "方块颜色："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   250
         TabIndex        =   4
         Top             =   330
         Width           =   1800
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   9615
      Begin VB.CommandButton Command5 
         Caption         =   "确定"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3120
         TabIndex        =   17
         Top             =   3960
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "规则：使用方向键控制方块避开障碍，接触到小方块时会有一定几率加分。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   16
         Top             =   1200
         Width           =   7455
      End
      Begin VB.Label Label4 
         Caption         =   "制作者：陈锡安"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   5865
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "商店"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   27.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      TabIndex        =   12
      Top             =   3240
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
    If Dir("D:\Flying_Squre", vbDirectory) = "" Then MkDir ("D:\Flying_Squre\")
    If Dir("D:\Flying_Squre\Game_Data") = "" Then
        Open "D:\Flying_Squre\Game_Data" For Output As #1
        Print #1, 1
        Close #1
    Else
        a = 0
        Open "D:\Flying_Squre\Game_Data" For Input As #1
        Do While Not EOF(1)
            Line Input #1, Data(a)
            a = a + 1
        Loop
        Close #1
        Open "D:\Flying_Squre\Game_Data" For Output As #2
        Print #2, 1
        Print #2, Data(1)
        Print #2, Data(2)
        Print #2, Data(3)
        Print #2, Data(4)
        Close #2
    End If
ElseIf Check1.Value = 0 Then
    If Not Dir("D:\Flying_Squre\Game_Data") = "" Then
        a = 0
        Open "D:\Flying_Squre\Game_Data" For Input As #1
        Do While Not EOF(1)
            Line Input #1, Data(a)
            a = a + 1
        Loop
        Close #1
        Open "D:\Flying_Squre\Game_Data" For Output As #2
        Print #2, 0
        Print #2, Data(1)
        Print #2, Data(2)
        Print #2, Data(3)
        Print #2, Data(4)
        Close #2
    End If
    If MsgBox("是否删除硬盘中的游戏信息文件", 36) = 6 And Not Dir("D:\Flying_Squre\Game_Data") = "" Then
        Kill "D:\Flying_Squre\Game_Data"
    End If
End If
End Sub
'Edited By Aden An Chen
'Apr 23,2017

Private Sub Command1_Click()
Form1.Label1 = ""
Form1.Timer1.Enabled = False
Form1.Timer2.Enabled = False
Form1.Timer3.Enabled = False
Form1.Timer5.Enabled = True
Count_Down = 0
Form1.Label1.Visible = True
Form1.Shape3.Visible = False
Form1.Show
Form2.Hide
End Sub

Private Sub Command2_Click()
Frame1.Visible = True
End Sub

Private Sub Command3_Click()
Color_Of_Squre = Color_Squre.BackColor
Color_Of_Obstacle = Color_Obstacle.BackColor
If Grade_Easy.Value = True Then Grade = 0
If Grade_Mid.Value = True Then Grade = 1
If Grade_Hard.Value = True Then Grade = 2
If Grade_Very_Hard.Value = True Then Grade = 3
Frame1.Visible = False
End Sub


Private Sub Command4_Click()
Form2.Hide
Form4.Show
End Sub

Private Sub Command5_Click()
Frame2.Visible = False
End Sub

Private Sub Command6_Click()
File_of_Sound = Text1.Text
If Not Right(File_of_Sound, 1) = "\" Then File_of_Sound = File_of_Sound + "\"
If Dir(File_of_Sound & "bomb.wav") = "" Then Sound_Not_Find = Sound_Not_Sound & vbCrLf & "bomb.wav"
If Dir(File_of_Sound & "buffer.wav") = "" Then Sound_Not_Find = Sound_Not_Find & vbCrLf & "buffer.wav"
If Dir(File_of_Sound & "wall_Broken.wav") = "" Then Sound_Not_Find = Sound_Not_Find & vbCrLf & "wall_Broken.wav"
If Dir(File_of_Sound & "wall_Mechine_Start.wav") = "" Then Sound_Not_Find = Sound_Not_Find & vbCrLf & "wall_Mechine_Start.wav"
If Not Sound_Not_Find = "" Then MsgBox ("在设定文件夹中仍有下列文件未找到：" & vbCrLf & Sound_Not_Find)
End Sub

Private Sub Form_Load()
If Not Dir("D:\Flying_Squre", vbDirectory) = "" Then
    If Not Dir("D:\Flying_Squre\Game_Data") = "" Then
        Open "D:\Flying_Squre\Game_Data" For Input As #1
        Line Input #1, used_or_not
        Close #1
        If used_or_not = 1 Then
            a = 0
            Open "D:\Flying_Squre\Game_Data" For Input As #1
            Do While Not EOF(1)
                Line Input #1, Data(a)
                a = a + 1
            Loop
            Close #1
            Check1.Value = Data(0)
            s2 = Data(1)
            Color_Of_Squre = Data(2)
            Color_Of_Obstacle = Data(3)
            Color_Squre.BackColor = Data(2)
            Color_Obstacle.BackColor = Data(3)
            File_of_Sound = Data(4)
        End If
    End If
End If
If File_of_Sound = "" Then File_of_Sound = "D:\Flying_Squre\"
If Dir(File_of_Sound & "bomb.wav") = "" Or Dir(File_of_Sound & "buffer.wav") = "" _
   Or Dir(File_of_Sound & "wall_Broken.wav") = "" Or Dir(File_of_Sound & "wall_Mechine_Start.wav") = "" Then
    a = MsgBox("在默认文件夹 " & File_of_Sound & "中没有发现所有游戏声音文件。" _
                           & vbCrLf & "请把所有游戏声音文件放在默认文件夹。" _
                           & vbCrLf & "或在设置中改变中把默认文件夹改为含有声音文件的文件夹。", vbOKOnly + vbInformation, "提示")
End If
End Sub
'Edited By Aden An Chen
'Apr 23,2017

Private Sub Form_Unload(Cancel As Integer)
Call Unload_All
End Sub

Private Sub List1_Click()
If List1.ListIndex = 0 Then Color_Squre.BackColor = vbBlack
If List1.ListIndex = 1 Then Color_Squre.BackColor = vbBlue
If List1.ListIndex = 2 Then Color_Squre.BackColor = vbGreen
If List1.ListIndex = 3 Then Color_Squre.BackColor = vbRed
If List1.ListIndex = 4 Then Color_Squre.BackColor = vbYellow
End Sub

Private Sub List2_Click()
If List2.ListIndex = 0 Then Color_Obstacle.BackColor = vbBlack
If List2.ListIndex = 1 Then Color_Obstacle.BackColor = vbBlue
If List2.ListIndex = 2 Then Color_Obstacle.BackColor = vbGreen
If List2.ListIndex = 3 Then Color_Obstacle.BackColor = vbRed
If List2.ListIndex = 4 Then Color_Obstacle.BackColor = vbYellow
End Sub
'Edited By Aden An Chen
'2016/8/21
Private Sub More_Color_Click(Index As Integer)
Color_Form.Show
Color_That_Change = Index
End Sub
'Edited By Aden An Chen
'Apr 23,2017
