VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "һԪ���η��̼�����"
   ClientHeight    =   8370
   ClientLeft      =   1785
   ClientTop       =   1530
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   15315
   Begin VB.Frame Frame1 
      Caption         =   "������ɫ"
      Height          =   2160
      Left            =   270
      TabIndex        =   21
      Top             =   5205
      Width           =   6075
      Begin VB.CommandButton Command5 
         Caption         =   "ѡ����ɫ"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4200
         TabIndex        =   27
         Top             =   915
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ȷ��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   26
         Top             =   1425
         Width           =   1950
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1665
         MaxLength       =   7
         TabIndex        =   25
         Text            =   "#"
         Top             =   855
         Width           =   2490
      End
      Begin VB.OptionButton Option2 
         Caption         =   "�Զ�����ɫ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3060
         TabIndex        =   23
         Top             =   240
         Width           =   2190
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�����ɫ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   210
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   1845
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "RGB��ɫ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         TabIndex        =   24
         Top             =   945
         Width           =   1620
      End
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "����"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4380
      TabIndex        =   18
      Top             =   7545
      Width           =   2115
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "����"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1320
      TabIndex        =   17
      Top             =   7560
      Width           =   2115
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   720
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2880
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   4920
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ͼ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   4440
      Width           =   1500
   End
   Begin VB.PictureBox P1 
      Height          =   8000
      Left            =   6630
      ScaleHeight     =   7935
      ScaleWidth      =   8445
      TabIndex        =   2
      Top             =   120
      Width           =   8500
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   0
      Top             =   4440
      Width           =   1500
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Y :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3495
      TabIndex        =   20
      Top             =   7530
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "X :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   375
      TabIndex        =   19
      Top             =   7500
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "X1="
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   360
      TabIndex        =   16
      Top             =   2505
      Width           =   810
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "X2="
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3105
      TabIndex        =   15
      Top             =   2520
      Width           =   810
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "��="
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   465
      TabIndex        =   14
      Top             =   3570
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "A��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   0
      TabIndex        =   13
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "B��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2040
      TabIndex        =   12
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "C��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4200
      TabIndex        =   11
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "AX^2+BX+C=0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   2640
   End
   Begin VB.Menu men_set 
      Caption         =   "����"
      Begin VB.Menu men_elean 
         Caption         =   "�Զ��������"
         Begin VB.Menu men_clean_open 
            Caption         =   "����"
         End
         Begin VB.Menu men_clean_off 
            Caption         =   "�ر�"
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rgb1 As Integer
Public rgb2 As Integer
Public rgb3 As Integer
Private Sub Command1_Click()
If Text1(0) = "" Then
   Text1(0) = "0"
End If
If Text1(1) = "" Then
   Text1(1) = "0"
End If
If Text1(2) = "" Then
   Text1(2) = "0"
End If
Text6 = Text1(1) ^ 2 - 4 * Text1(0) * Text1(2)
If Text6 < 0 Then
   X = MsgBox("�˷����޽�", 48, "��ʾ")
Else
   Text4 = (Val(-Text2) + Sqr(Text6)) / 2 * Text1(0)
   Text5 = (Val(-Text2) - Sqr(Text6)) / 2 * Text1(0)
End If
End Sub

Private Sub Command4_Click()
If Len(Text9.Text) = 7 Then
    rgb1 = Val("&H" & Mid(Text9.Text, 1, 2))
    rgb2 = Val("&H" & Mid(Text9.Text, 4, 2))
    rgb3 = Val("&H" & Right(Text9.Text, 2))
    co = RGB(rgb1, rgb2, rgb3)
End If
End Sub

Private Sub Command5_Click()
Form3.Show
End Sub

Private Sub Form_Load()
P1.Scale (-10, 10)-(10, -10)
End Sub

Private Sub Command3_Click()
P1.Picture = LoadPicture("")
Text1(0) = ""
Text1(1) = ""
Text1(2) = ""
Text4 = ""
Text5 = ""
Text6 = ""
End Sub


Private Sub Command2_Click()
If Text1(0) = "" Then
   Text1(0) = "0"
End If
If Text1(1) = "" Then
   Text1(1) = "0"
End If
If Text1(2) = "" Then
   Text1(2) = "0"
End If
Dim X As Single
P1.Line (-10, 0)-(10, 0)
P1.Line (0, -10)-(0, 10)
If Option1.Value = True Then
    co = RGB(Rnd * 256, Rnd * 256, Rnd * 256)
End If
For X = -10 To 10 Step 0.01
    P1.PSet (X, Text1(0) * X ^ 2 + Text1(1) * X + Text1(2)), co
Next X
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form3
End Sub

Private Sub men_clean_off_Click()
men_clean_off.Enabled = False
men_clean_open.Enabled = True
End Sub

Private Sub men_clean_open_Click()
men_clean_open.Enabled = False
men_clean_off.Enabled = True
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then Command4.Enabled = False: Command5.Enabled = False
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then Command4.Enabled = True: Command5.Enabled = True
End Sub

Private Sub P1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
P1.MousePointer = 2
Text7.Text = Int(X * 1000) / 1000
Text8.Text = Int(Y * 1000) / 1000
End Sub

Private Sub Text1_Click(Index As Integer)
If men_clean_open.Enabled = False Then Text1(Index).Text = ""
If Text1(Index).Text = "picture" Then Form1.Show: Form1.Timer1.Enabled = False
End Sub

Private Sub Text9_Click()
If men_clean_open.Enabled = False Then Text9.Text = "#"
End Sub
'Edited by Aden
'2016/9/24
