VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1440
      Top             =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "硬盘正在高级格式化"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   30
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   5400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Ctrl As Integer
Public b As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 1 Then
    If Ctrl = 0 Then
        Ctrl = 1
    Else
        If Ctrl = 1 Then
            Ctrl = 2
        Else
            If Ctrl = 2 Then Unload Form1
        End If
    End If
Else
    Ctrl = 0
End If
End Sub

Private Sub Timer1_Timer()
b = b + 1
a = ""
For x = 0 To 227 - Rnd * 226
    a = a & Chr(Int(Rnd * (130 - 50 + 1) + 50))
Next x
Print a
If b >= 64 Then
    Timer1.Enabled = False
    TimeDelay (5 + Rnd)
    Form1.Cls
    Label1.Visible = True
    b = 1
    Label1.Caption = Label1.Caption & vbCrLf & vbCrLf & Str(b) & "%"
    For x = 1 To 99
        b = b + 1
        TimeDelay (Rnd * 2)
        Label1.Caption = "硬盘正在高级格式化" & vbCrLf & vbCrLf & Str(b) & "%"
    Next x
    TimeDelay (2 + Rnd)
    Label1.Caption = "正在建立系统后门"
    TimeDelay (2 + Rnd)
    Label1.Caption = "建立完成"
    TimeDelay (1)
    b = 0
    For x = 1 To 100
        TimeDelay (Rnd * 10)
        b = b + 1
        Label1.Caption = "硬盘出现错误,无法正常打开Windows" & vbCrLf & "Windows正在修复所有错误" & vbCrLf & vbCrLf & Str(b) & "%"
    Next x
    Label1.Caption = "修复完成,请手动重新启动"
End If
End Sub

Public Sub TimeDelay(ByVal PauseSecond As Single)
 Dim Star, PauseTime
 Star = Timer
 PauseTime = PauseSecond
 Do While Timer < Star + PauseTime
 DoEvents
 Loop
End Sub

