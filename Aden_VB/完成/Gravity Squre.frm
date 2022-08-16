VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gravity Squre"
   ClientHeight    =   7560
   ClientLeft      =   4635
   ClientTop       =   1980
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   13905
   Begin VB.Frame Frame1 
      Height          =   7575
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   13815
      Begin VB.Frame Frame2 
         Height          =   7455
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   13815
         Begin VB.CommandButton Command3 
            Caption         =   "Play"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   26.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   5160
            TabIndex        =   8
            Top             =   3000
            Width           =   3375
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   8160
         TabIndex        =   6
         Top             =   4320
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Play Again"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3000
         TabIndex        =   5
         Top             =   4320
         Width           =   3375
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5880
         TabIndex        =   4
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Your Score Is :"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   960
         TabIndex        =   3
         Top             =   1440
         Width           =   4050
      End
   End
   Begin VB.Timer Timer8 
      Interval        =   5000
      Left            =   6720
      Top             =   3960
   End
   Begin VB.Timer Timer7 
      Interval        =   100
      Left            =   4320
      Top             =   3840
   End
   Begin VB.Timer Timer6 
      Interval        =   1500
      Left            =   6720
      Top             =   3600
   End
   Begin VB.Timer Timer5 
      Interval        =   200
      Left            =   4320
      Top             =   3360
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4680
      Top             =   6960
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   9360
      Top             =   3000
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   2880
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4320
      Top             =   120
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   9840
      Shape           =   3  'Circle
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   10920
      Shape           =   3  'Circle
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   12360
      Shape           =   3  'Circle
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   10200
      Shape           =   3  'Circle
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10440
      TabIndex        =   1
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Your Score Is :"
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
      Left            =   10080
      TabIndex        =   0
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   9840
      X2              =   9840
      Y1              =   0
      Y2              =   7560
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4560
      Top             =   4320
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Sub Command1_Click()
Shape1.Top = 4320
Shape1.Left = 4320
Shape3.Top = 3000
Shape3.Left = 4000
co = 0
c1 = 0
n = 0
s = 0
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Timer7.Enabled = True
Timer8.Enabled = True
Frame1.Visible = False
End Sub

Private Sub Command2_Click()
Shape1.Top = 4320
Shape1.Left = 4320
Shape3.Top = 3000
Shape3.Left = 4000
co = 0
c1 = 0
n = 0
s = 0
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Timer7.Enabled = True
Timer8.Enabled = True
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False
Frame2.Visible = True
End Sub

Private Sub Command3_Click()
Shape1.Top = 4320
Shape1.Left = 4320
Shape3.Top = 3000
Shape3.Left = 4000
co = 0
c1 = 0
n = 0
s = 0
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Timer7.Enabled = True
Timer8.Enabled = True
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False
Frame1.Visible = False
Frame2.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
   co = 1
ElseIf KeyCode = 38 Then
   co = 2
ElseIf KeyCode = 39 Then
   co = 3
ElseIf KeyCode = 40 Then
   co = 4
End If
End Sub

Private Sub Form_Load()
Frame1.Visible = True
Frame2.Visible = True
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
End Sub

Private Sub Timer1_Timer()
If c1 = 1 Then

Else
 Shape1.Top = Shape1.Top - 150
End If
End Sub

Private Sub Timer2_Timer()
If c1 = 1 Then

Else
   Shape1.Left = Shape1.Left - 150
End If
End Sub

Private Sub Timer3_Timer()
If c1 = 1 Then

Else
   Shape1.Left = Shape1.Left + 150
End If
End Sub

Private Sub Timer4_Timer()
If c1 = 1 Then

Else
   Shape1.Top = Shape1.Top + 150
End If
End Sub

Private Sub Timer5_Timer()
If co = 0 Then
   Timer1.Enabled = False
   Timer2.Enabled = False
   Timer3.Enabled = False
   Timer4.Enabled = False
ElseIf co = 1 Then
   Timer1.Enabled = False
   Timer2.Enabled = True
   Timer3.Enabled = False
   Timer4.Enabled = False
ElseIf co = 2 Then
   Timer1.Enabled = True
   Timer2.Enabled = False
   Timer3.Enabled = False
   Timer4.Enabled = False
ElseIf co = 3 Then
   Timer1.Enabled = False
   Timer2.Enabled = False
   Timer3.Enabled = True
   Timer4.Enabled = False
ElseIf co = 4 Then
   Timer1.Enabled = False
   Timer2.Enabled = False
   Timer3.Enabled = False
   Timer4.Enabled = True
End If
If IsTouched(Shape1, Shape2) Then
   a(0) = Rndz(100, 6900)
   a(1) = Rndz(100, 9000)
   Shape2.Visible = False
   Shape2.Top = a(0)
   Shape2.Left = a(1)
   Shape2.Visible = True
   s = s + 1
   Label2.Caption = s
   Label4.Caption = s
End If
End Sub
Function IsTouched(a, b) As Boolean
IsTouched = Not ( _
    a.Left > b.Left + b.Width Or _
    b.Left > a.Left + a.Width Or _
    a.Top > b.Top + b.Height Or _
    b.Top > a.Top + a.Height)
End Function
Private Function Rndz(a As Long, b As Long)
    Randomize
    Rndz = Int((a - b + 1) * Rnd() + b)
End Function


Private Sub Timer6_Timer()
b(0) = Rndz(100, 6900)
b(1) = Rndz(100, 9000)
Shape3.Visible = False
Shape3.Top = b(0)
Shape3.Left = b(1)
Shape3.Visible = True
End Sub

Private Sub Timer7_Timer()

If IsTouched(Shape1, Shape3) Then
co = 0: Sleep 500: Frame1.Visible = True: Timer5.Enabled = False
Timer5.Enabled = False: Timer6.Enabled = False: Timer7.Enabled = False: Timer8.Enabled = False: c1 = 1
ElseIf IsTouched(Shape1, Shape4) Then
co = 0: Sleep 500: Frame1.Visible = True: Timer5.Enabled = False
Timer5.Enabled = False: Timer6.Enabled = False: Timer7.Enabled = False: Timer8.Enabled = False: c1 = 1
ElseIf IsTouched(Shape1, Shape5) Then
co = 0: Sleep 500: Frame1.Visible = True: Timer5.Enabled = False
Timer5.Enabled = False: Timer6.Enabled = False: Timer7.Enabled = False: Timer8.Enabled = False: c1 = 1
ElseIf IsTouched(Shape1, Shape6) Then
co = 0: Sleep 500: Frame1.Visible = True: Timer5.Enabled = False
Timer5.Enabled = False: Timer6.Enabled = False: Timer7.Enabled = False: Timer8.Enabled = False: c1 = 1
ElseIf IsTouched(Shape1, Shape7) Then
co = 0: Sleep 500: Frame1.Visible = True: Timer5.Enabled = False
Timer5.Enabled = False: Timer6.Enabled = False: Timer7.Enabled = False: Timer8.Enabled = False: c1 = 1
ElseIf Shape1.Top <= 100 Or Shape1.Top >= 6900 Or Shape1.Left <= 100 Or Shape1.Left >= 9100 Then
co = 0: Sleep 500: Frame1.Visible = True: Timer5.Enabled = False
Timer5.Enabled = False: Timer6.Enabled = False: Timer7.Enabled = False: Timer8.Enabled = False: c1 = 1
End If
End Sub

Private Sub Timer8_Timer()
n = n + 1
If n = 1 Then
   Shape4.Top = 5783
   Shape4.Left = 2844
   Shape4.Visible = True
ElseIf n = 2 Then
   Shape5.Top = 2878
   Shape5.Left = 232
   Shape5.Visible = True
   Shape4.Height = Shape4.Height + 100
   Shape4.Width = Shape4.Width + 100
ElseIf n = 3 Then
   Shape6.Top = 6823
   Shape6.Left = 8732
   Shape6.Visible = True
   Shape4.Height = Shape4.Height + 100
   Shape4.Width = Shape4.Width + 100
   Shape5.Height = Shape4.Height + 100
   Shape5.Width = Shape4.Width + 100
ElseIf n = 4 Then
   Shape7.Top = 927
   Shape7.Left = 5678
   Shape7.Visible = True
   Shape4.Height = Shape4.Height + 100
   Shape4.Width = Shape4.Width + 100
   Shape4.Height = Shape5.Height + 100
   Shape4.Width = Shape5.Width + 100
   Shape4.Height = Shape6.Height + 100
   Shape4.Width = Shape6.Width + 100
ElseIf n >= 5 Then
   Shape4.Height = Shape4.Height + 100
   Shape4.Width = Shape4.Width + 100
   Shape4.Height = Shape5.Height + 100
   Shape4.Width = Shape5.Width + 100
   Shape4.Height = Shape6.Height + 100
   Shape4.Width = Shape6.Width + 100
   Shape4.Height = Shape7.Height + 100
   Shape4.Width = Shape7.Width + 100
End If
End Sub
'**********************************************END********************************************
'Edited by Aden
'2016/8/21

