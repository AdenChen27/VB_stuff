VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "查找"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   12105
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5010
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1170
      Width           =   11955
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   105
      TabIndex        =   1
      Top             =   30
      Width           =   11910
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查找"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11955
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim txt As String
Dim a As Integer
Dim b
Open "C:\Users\zhchen\Desktop\Aden_VB\test.txt" For Input As #1
On Error Resume Next
    Dim key As String
    Dim value As String
    Do While a < 100
        a = a + 1
        If b = 0 Then b = 1 Else: b = 0
        Input #1, key
        If b = 1 Then
            If key = "1" & Text1.Text Then Exit Do: Input #1, value: Text2.Text = value
        End If
    Loop
    Close #1
End Sub
