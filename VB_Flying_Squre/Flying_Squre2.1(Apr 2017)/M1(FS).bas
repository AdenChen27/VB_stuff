Attribute VB_Name = "Module1"
Public Color_Of_Obstacle As Long
Public Color_Of_Squre As Long
Public Count_Down As Integer
Public Grade As Integer
Public p1 As Integer
Public t1 As Integer 'speed
Public s1 As Integer 'score got in last game
Public s2 As Integer 'total score
Public sh(1) As Integer 'num of the buyings
Public of(1) As Integer
Public sa As Integer
Public Color_That_Change As Integer
Public File_of_Sound As String
Public Data(4) As String

Public Sub TimeDelay(ByVal PauseSecond As Single)
 Dim Star, PauseTime
 Star = Timer
 PauseTime = PauseSecond
 Do While Timer < Star + PauseTime
 DoEvents
 Loop
End Sub

Public Sub Unload_All()
Unload Form1
Unload Form2
Unload Form3
Unload Form4
Unload Color_Form
If Not Dir("D:\Flying_Squre\Game_Data") = "" Then
    Open "D:\Flying_Squre\Game_Data" For Input As #1
    Line Input #1, used_or_not
    Close #1
    If used_or_not = 1 Then
        Open "D:\Flying_Squre\Game_Data" For Output As #2
        Print #2, 1
        Print #2, s2
        Print #2, Color_Of_Squre
        Print #2, Color_Of_Obstacle
        Print #2, File_of_Sound
        Close #2
    End If
    Close #1
End If
End Sub
'Edited By Aden An Chen
'Apr 23,2017
