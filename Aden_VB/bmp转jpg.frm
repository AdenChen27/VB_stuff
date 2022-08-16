VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   14025
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1710
      Left            =   13485
      TabIndex        =   1
      Top             =   1770
      Width           =   555
   End
   Begin VB.PictureBox Picture1 
      Height          =   7035
      Left            =   75
      ScaleHeight     =   6975
      ScaleWidth      =   13260
      TabIndex        =   0
      Top             =   210
      Width           =   13320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    type As Long
    Value As Long
End Type

Private Type EncoderParameters
    Count As Long
    Parameter As EncoderParameter
End Type

Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hpal As Long, Bitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As Long, Bitmap As Long) As Long

Private Sub Command1_Click()
    Dim ret As Boolean
      
    Picture1.Picture = LoadPicture("C:\C++\About_C++\P 1.bmp") '打开要压缩的图片
      
    ret = PictureBoxSaveJPG(Picture1, "C:\C++\About_C++jpg 1.jpg") '保存压缩后的图片
    If ret = False Then
        MsgBox "保存失败"
    End If
End Sub

Private Function PictureBoxSaveJPG(ByVal pict As StdPicture, ByVal filename As String, Optional ByVal quality As Byte = 80) As Boolean
    Dim tSI As GdiplusStartupInput
    Dim lRes As Long
    Dim lGDIP As Long
    Dim lBitmap As Long
    tSI.GdiplusVersion = 1
    lRes = GdiplusStartup(lGDIP, tSI, 0)
    If lRes = 0 Then
        lRes = GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap)
        If lRes = 0 Then
            Dim tJpgEncoder As GUID
            Dim tParams As EncoderParameters
            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
            tParams.Count = 1
            With tParams.Parameter ' Quality
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
                .NumberOfValues = 1
                .type = 4
                .Value = VarPtr(quality)
            End With
            lRes = GdipSaveImageToFile(lBitmap, StrPtr(filename), tJpgEncoder, tParams)
            GdipDisposeImage lBitmap
        End If
        GdiplusShutdown lGDIP
    End If
    If lRes Then
        PictureBoxSaveJPG = False
    Else
        PictureBoxSaveJPG = True
    End If
End Function
