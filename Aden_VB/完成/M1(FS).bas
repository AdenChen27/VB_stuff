Attribute VB_Name = "Module1"
Public a2 As Integer
Public a3 As Integer
Public a4 As Integer
Public a5 As Integer
Public p1 As Integer
Public t1 As Integer
Public s1 As Integer
Public s2 As Integer
Public sh(1) As Integer
Public of(1) As Integer
Public sa As Integer

Public Sub TimeDelay(ByVal PauseSecond As Single)
 Dim Star, PauseTime
 Star = Timer
 PauseTime = PauseSecond
 Do While Timer < Star + PauseTime
 DoEvents
 Loop
End Sub

