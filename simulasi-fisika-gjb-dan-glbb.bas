Attribute VB_Name = "Module1"
Private Sub Button2_Click()
 Range("D4").Value = 0 '0
  While Range("D4") < 10
  Range("D4").Value = Range("D4") + 0.05
  DoEvents
 Wend
 
End Sub
Private Sub Button4_Click()
 Range("J5").Value = 0 '0
  While Range("J5") < 10
  Range("J5").Value = Range("J5") + 0.05
  DoEvents
 Wend
End Sub

